using ManagedShell.Common.Helpers;
using ManagedShell.Common.Logging;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Forms;
using ManagedShell.Common.Enums;
using ManagedShell.Common.SupportingClasses;
using static ManagedShell.Interop.NativeMethods;

namespace ManagedShell.WindowsTasks
{
    public class TasksService : DependencyObject, IDisposable
    {
        public static readonly IconSize DEFAULT_ICON_SIZE = IconSize.Small;

        public event EventHandler<WindowEventArgs> WindowActivated;
        public event EventHandler<EventArgs> DesktopActivated;
        public event EventHandler<FullScreenEventArgs> FullScreenChanged;
        public event EventHandler<WindowEventArgs> MonitorChanged;

        public Func<ApplicationWindow, IList<ApplicationWindow>, int> WindowInsertionIndexProvider { get; set; }

        // Off by default: some apps (observed with a Word add-in) briefly create and activate a
        // genuine but 0x0-sized top-level window as a side effect of their own internal handling,
        // which - like any other window that passes the other CanAddToTaskbar checks - gets a real
        // task button for the instant it exists. Host apps can opt into filtering these out via
        // ApplicationWindow.CanAddToTaskbar; left off by default since a window with no area is an
        // edge case some legitimate (if unusual) windows could conceivably hit.
        public bool FilterZeroSizeWindows { get; set; }

        private NativeWindowEx _HookWin;
        private object _windowsLock = new object();
        internal bool IsInitialized;
        private IconSize _taskIconSize;

        private static int WM_SHELLHOOKMESSAGE = -1;
        private static int WM_TASKBARCREATEDMESSAGE = -1;
        private static int TASKBARBUTTONCREATEDMESSAGE = -1;
        private static IntPtr cloakEventHook = IntPtr.Zero;
        private WinEventProc cloakEventProc;
        private static IntPtr moveEventHook = IntPtr.Zero;
        private WinEventProc moveEventProc;
        private static IntPtr desktopSwitchEventHook = IntPtr.Zero;
        private WinEventProc desktopSwitchEventProc;
        private System.Windows.Threading.DispatcherTimer _desktopSwitchScanTimer;

        internal ITaskCategoryProvider TaskCategoryProvider;
        private TaskCategoryChangeDelegate CategoryChangeDelegate;

        public IconSize TaskIconSize
        {
            get { return _taskIconSize; }
            set
            {
                if (value == _taskIconSize)
                {
                    return;
                }

                _taskIconSize = value;

                if (!IsInitialized)
                {
                    return;
                }

                foreach (var window in Windows)
                {
                    if (!window.ShowInTaskbar)
                    {
                        return;
                    }

                    window.UpdateProperties();
                }
            }
        }

        public TasksService() : this(DEFAULT_ICON_SIZE)
        {
        }
        
        public TasksService(IconSize iconSize)
        {
            TaskIconSize = iconSize;
        }

        internal void Initialize(bool withMultiMonTracking)
        {
            if (IsInitialized)
            {
                return;
            }

            try
            {
                ShellLogger.Debug("TasksService: Starting");

                // create window to receive task events
                _HookWin = new NativeWindowEx();
                _HookWin.CreateHandle(new CreateParams());

                // prevent other shells from working properly
                SetTaskmanWindow(_HookWin.Handle);

                // register to receive task events
                RegisterShellHookWindow(_HookWin.Handle);
                WM_SHELLHOOKMESSAGE = RegisterWindowMessage("SHELLHOOK");
                WM_TASKBARCREATEDMESSAGE = RegisterWindowMessage("TaskbarCreated");
                TASKBARBUTTONCREATEDMESSAGE = RegisterWindowMessage("TaskbarButtonCreated");
                _HookWin.MessageReceived += ShellWinProc;

                if (EnvironmentHelper.IsWindows8OrBetter)
                {
                    // set event hook for cloak/uncloak events
                    cloakEventProc = CloakEventCallback;

                    if (cloakEventHook == IntPtr.Zero)
                    {
                        cloakEventHook = SetWinEventHook(
                            EVENT_OBJECT_CLOAKED,
                            EVENT_OBJECT_UNCLOAKED,
                            IntPtr.Zero,
                            cloakEventProc,
                            0,
                            0,
                            WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS);
                    }

                    // Hook EVENT_SYSTEM_DESKTOPSWITCH for non-virtual desktop switches
                    // (UAC prompt, lock screen, fast-user-switch). Windows 10/11 virtual desktop
                    // switches do NOT fire this event — they only generate per-window
                    // EVENT_OBJECT_CLOAKED/UNCLOAKED events, which CloakEventCallback handles.
                    desktopSwitchEventProc = DesktopSwitchEventCallback;

                    if (desktopSwitchEventHook == IntPtr.Zero)
                    {
                        desktopSwitchEventHook = SetWinEventHook(
                            EVENT_SYSTEM_DESKTOPSWITCH,
                            EVENT_SYSTEM_DESKTOPSWITCH,
                            IntPtr.Zero,
                            desktopSwitchEventProc,
                            0,
                            0,
                            WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS);
                    }
                }

                if (withMultiMonTracking && !EnvironmentHelper.IsWindows8OrBetter)
                {
                    // set event hook for move events
                    // In Windows 8 and newer, use HSHELL_MONITORCHANGED instead
                    moveEventProc = MoveEventCallback;

                    if (moveEventHook == IntPtr.Zero)
                    {
                        moveEventHook = SetWinEventHook(
                            EVENT_OBJECT_LOCATIONCHANGE,
                            EVENT_OBJECT_LOCATIONCHANGE,
                            IntPtr.Zero,
                            moveEventProc,
                            0,
                            0,
                            WINEVENT_OUTOFCONTEXT);
                    }
                }

                // set window for ITaskbarList
                setTaskbarListHwnd(_HookWin.Handle);

                // adjust minimize animation
                SetMinimizedMetrics();

                // enumerate windows already opened and set active window
                getInitialWindows();

                IsInitialized = true;
            }
            catch (Exception ex)
            {
                ShellLogger.Info("TasksService: Unable to start: " + ex.Message);
            }
        }

        internal void SetTaskCategoryProvider(ITaskCategoryProvider provider)
        {
            TaskCategoryProvider = provider;

            if (CategoryChangeDelegate == null)
            {
                CategoryChangeDelegate = CategoriesChanged;
            }

            TaskCategoryProvider.SetCategoryChangeDelegate(CategoryChangeDelegate);
        }

        private void getInitialWindows()
        {
            EnumWindows((hwnd, lParam) =>
            {
                ApplicationWindow win = new ApplicationWindow(this, hwnd);

                if (win.CanAddToTaskbar && win.ShowInTaskbar && !Windows.Contains(win))
                {
                    Windows.Add(win);

                    sendTaskbarButtonCreatedMessage(win.Handle);
                }

                return true;
            }, 0);

            IntPtr hWndForeground = GetForegroundWindow();
            if (Windows.Any(i => i.Handle == hWndForeground && i.ShowInTaskbar))
            {
                ApplicationWindow win = Windows.First(wnd => wnd.Handle == hWndForeground);
                win.State = ApplicationWindow.WindowState.Active;
                win.SetShowInTaskbar();
            }
        }

        public void Dispose()
        {
            if (IsInitialized)
            {
                ShellLogger.Debug("TasksService: Deregistering hooks");
                DeregisterShellHookWindow(_HookWin.Handle);
                if (cloakEventHook != IntPtr.Zero) { UnhookWinEvent(cloakEventHook); cloakEventHook = IntPtr.Zero; }
                if (moveEventHook != IntPtr.Zero) { UnhookWinEvent(moveEventHook); moveEventHook = IntPtr.Zero; }
                if (desktopSwitchEventHook != IntPtr.Zero) { UnhookWinEvent(desktopSwitchEventHook); desktopSwitchEventHook = IntPtr.Zero; }
                _HookWin.DestroyHandle();
                setTaskbarListHwnd(IntPtr.Zero);
                IsInitialized = false;
                Windows.Clear();
            }

            TaskCategoryProvider?.Dispose();
        }

        private void CategoriesChanged()
        {
            foreach (ApplicationWindow window in Windows)
            {
                if (window.ShowInTaskbar)
                {
                    window.Category = TaskCategoryProvider?.GetCategory(window);
                }
            }
        }

        private void SetMinimizedMetrics()
        {
            MinimizedMetrics mm = new MinimizedMetrics
            {
                cbSize = (uint)Marshal.SizeOf(typeof(MinimizedMetrics))
            };

            IntPtr mmPtr = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(MinimizedMetrics)));

            try
            {
                Marshal.StructureToPtr(mm, mmPtr, true);
                SystemParametersInfo(SPI.GETMINIMIZEDMETRICS, mm.cbSize, mmPtr, SPIF.None);
                mm.iWidth = 140;
                mm.iArrange |= MinimizedMetricsArrangement.Hide;
                Marshal.StructureToPtr(mm, mmPtr, true);
                SystemParametersInfo(SPI.SETMINIMIZEDMETRICS, mm.cbSize, mmPtr, SPIF.None);
            }
            finally
            {
                Marshal.DestroyStructure(mmPtr, typeof(MinimizedMetrics));
                Marshal.FreeHGlobal(mmPtr);
            }
        }

        public void CloseWindow(ApplicationWindow window)
        {
            if (window.DoClose() != IntPtr.Zero)
            {
                ShellLogger.Debug($"TasksService: Removing window {window.Title} from collection due to no response");
                window.Dispose();
                Windows.Remove(window);
            }
        }

        private void sendTaskbarButtonCreatedMessage(IntPtr hWnd)
        {
            // Server Core doesn't support ITaskbarList, so sending this message on that OS could cause some assuming apps to crash
            if (!EnvironmentHelper.IsServerCore) SendNotifyMessage(hWnd, (uint)TASKBARBUTTONCREATEDMESSAGE, UIntPtr.Zero, IntPtr.Zero);
        }

        private ApplicationWindow addWindow(IntPtr hWnd, ApplicationWindow.WindowState initialState = ApplicationWindow.WindowState.Inactive, bool sanityCheck = false)
        {
            ApplicationWindow win = new ApplicationWindow(this, hWnd);

            // set window state if a non-default value is provided
            if (initialState != ApplicationWindow.WindowState.Inactive) win.State = initialState;

            // add window unless we need to validate it is eligible to show in taskbar
            if (!sanityCheck || win.CanAddToTaskbar)
            {
                int insertIdx = WindowInsertionIndexProvider?.Invoke(win, Windows) ?? -1;
                if (insertIdx >= 0 && insertIdx < Windows.Count)
                    Windows.Insert(insertIdx, win);
                else
                    Windows.Add(win);
                ShellLogger.Debug($"TasksService: Added window {hWnd} ({win.Title})");
            }

            // Only send TaskbarButtonCreated if we are shell, and if OS is not Server Core
            // This is because if Explorer is running, it will send the message, so we don't need to
            if (EnvironmentHelper.IsAppRunningAsShell) sendTaskbarButtonCreatedMessage(win.Handle);

            return win;
        }

        private void removeWindow(IntPtr hWnd)
        {
            if (Windows.Any(i => i.Handle == hWnd))
            {
                do
                {
                    ApplicationWindow win = Windows.First(wnd => wnd.Handle == hWnd);
                    win.Dispose();
                    Windows.Remove(win);

                    ShellLogger.Debug($"TasksService: Removed window {hWnd} ({win.Title})");
                }
                while (Windows.Any(i => i.Handle == hWnd));
            }
        }

        // Manual escape hatch for windows stuck in the collection with a dead hwnd. Normally
        // removeWindow() only runs off the shell hook's HSHELL_WINDOWDESTROYED notification (see
        // ShellWinProc below); if that notification is ever missed - the process was killed rather
        // than closed gracefully, or the message was simply dropped - the ApplicationWindow lingers
        // forever, since nothing else ever re-validates existing entries (ScanForNewWindows only
        // adds windows it hasn't seen yet, it never prunes ones it has). Returns the number removed.
        public int SweepDeadWindows()
        {
            int removed = 0;
            lock (_windowsLock)
            {
                foreach (var hwnd in Windows.Select(w => w.Handle).ToList())
                {
                    if (!IsWindow(hwnd))
                    {
                        removeWindow(hwnd);
                        removed++;
                    }
                }
            }

            ShellLogger.Info($"TasksService: SweepDeadWindows removed {removed} dead window(s)");
            return removed;
        }

        private void redrawWindow(ApplicationWindow win)
        {
            win.UpdateProperties();
            ShellLogger.Debug($"TasksService: Updated window {win.Handle} ({win.Title})");

            foreach (ApplicationWindow wind in Windows)
            {
                if (wind.WinFileName == win.WinFileName && wind.Handle != win.Handle)
                {
                    wind.UpdateProperties();
                }
            }
        }

        private void ShellWinProc(ref Message msg, ref bool handled)
        {
            Message msgCopy = msg;
            handled = true;
            if (msg.Msg == WM_SHELLHOOKMESSAGE)
            {
                try
                {
                    lock (_windowsLock)
                    {
                        switch ((HSHELL)msg.WParam.ToInt32())
                        {
                            case HSHELL.WINDOWCREATED:
                                if (!Windows.Any(i => i.Handle == msgCopy.LParam))
                                {
                                    addWindow(msg.LParam);
                                }
                                else
                                {
                                    ApplicationWindow win = Windows.First(wnd => wnd.Handle == msgCopy.LParam);
                                    win.UpdateProperties();
                                }
                                break;

                            case HSHELL.WINDOWDESTROYED:
                                removeWindow(msg.LParam);
                                break;

                            case HSHELL.WINDOWREPLACING:
                                if (Windows.Any(i => i.Handle == msgCopy.LParam))
                                {
                                    ApplicationWindow win = Windows.First(wnd => wnd.Handle == msgCopy.LParam);
                                    win.State = ApplicationWindow.WindowState.Inactive;
                                    win.SetShowInTaskbar();
                                }
                                else
                                {
                                    addWindow(msg.LParam);
                                }
                                break;
                            case HSHELL.WINDOWREPLACED:
                                // TODO: If a window gets replaced, we lose app-level state such as overlay icons.
                                removeWindow(msg.LParam);
                                break;

                            case HSHELL.WINDOWACTIVATED:
                            case HSHELL.RUDEAPPACTIVATED:
                                // Ignore activation of windows that aren't real taskbar-eligible app windows
                                // and aren't already tracked (e.g. a WS_EX_NOACTIVATE shell window forcing
                                // itself to the foreground, as RetroBar does to keep a context menu on top).
                                // That isn't a genuine app switch, so it shouldn't deactivate every currently
                                // active tracked window.
                                if (msg.LParam != IntPtr.Zero
                                    && !Windows.Any(i => i.Handle == msgCopy.LParam)
                                    && !new ApplicationWindow(this, msg.LParam).CanAddToTaskbar)
                                {
                                    break;
                                }

                                foreach (var aWin in Windows.Where(w => w.State == ApplicationWindow.WindowState.Active))
                                {
                                    aWin.State = ApplicationWindow.WindowState.Inactive;
                                }

                                if (msg.LParam != IntPtr.Zero)
                                {
                                    ApplicationWindow win = null;

                                    if (Windows.Any(i => i.Handle == msgCopy.LParam))
                                    {
                                        win = Windows.First(wnd => wnd.Handle == msgCopy.LParam);
                                        win.State = ApplicationWindow.WindowState.Active;
                                        win.SetShowInTaskbar();
                                        ShellLogger.Debug($"TasksService: Activated window {win.Handle} ({win.Title})");
                                    }
                                    else
                                    {
                                        win = addWindow(msg.LParam, ApplicationWindow.WindowState.Active);
                                    }

                                    if (win != null)
                                    {
                                        foreach (ApplicationWindow wind in Windows)
                                        {
                                            if (wind.WinFileName == win.WinFileName && wind.Handle != win.Handle)
                                                wind.SetShowInTaskbar();
                                        }

                                        WindowEventArgs args = new WindowEventArgs
                                        {
                                            Window = win
                                        };

                                        WindowActivated?.Invoke(this, args);
                                    }
                                }
                                else
                                {
                                    DesktopActivated?.Invoke(this, new EventArgs());
                                }
                                break;

                            case HSHELL.FLASH:
                                if (Windows.Any(i => i.Handle == msgCopy.LParam))
                                {
                                    ApplicationWindow win = Windows.First(wnd => wnd.Handle == msgCopy.LParam);
                                    
                                    if (win.State != ApplicationWindow.WindowState.Active)
                                    {
                                        win.State = ApplicationWindow.WindowState.Flashing;
                                    }

                                    redrawWindow(win);
                                }
                                else
                                {
                                    addWindow(msg.LParam, ApplicationWindow.WindowState.Flashing, true);
                                }
                                break;

                            case HSHELL.ACTIVATESHELLWINDOW:
                                ShellLogger.Debug("TasksService: Activate shell window called.");
                                break;

                            case HSHELL.ENDTASK:
                                removeWindow(msg.LParam);
                                break;

                            case HSHELL.REDRAW:
                                if (Windows.Any(i => i.Handle == msgCopy.LParam))
                                {
                                    ApplicationWindow win = Windows.First(wnd => wnd.Handle == msgCopy.LParam);

                                    if (win.State == ApplicationWindow.WindowState.Flashing)
                                    {
                                        win.State = ApplicationWindow.WindowState.Inactive;
                                    }

                                    redrawWindow(win);
                                }
                                else
                                {
                                    addWindow(msg.LParam, ApplicationWindow.WindowState.Inactive, true);
                                }
                                break;

                            case HSHELL.MONITORCHANGED:
                                if (Windows.Any(i => i.Handle == msgCopy.LParam))
                                {
                                    ApplicationWindow win = Windows.First(wnd => wnd.Handle == msgCopy.LParam);
                                    win.SetMonitor();

                                    WindowEventArgs args = new WindowEventArgs
                                    {
                                        Window = win
                                    };

                                    MonitorChanged?.Invoke(this, args);
                                }
                                break;

                            case HSHELL.FULLSCREENENTER:
                                {
                                    FullScreenEventArgs args = new FullScreenEventArgs
                                    {
                                        Handle = msgCopy.LParam,
                                        IsEntering = true
                                    };

                                    FullScreenChanged?.Invoke(this, args);
                                    ShellLogger.Debug($"TasksService: Full screen entered by window {msgCopy.LParam}");
                                    break;
                                }

                            case HSHELL.FULLSCREENEXIT:
                                {
                                    FullScreenEventArgs args = new FullScreenEventArgs
                                    {
                                        Handle = msgCopy.LParam,
                                        IsEntering = false
                                    };

                                    FullScreenChanged?.Invoke(this, args);
                                    ShellLogger.Debug($"TasksService: Full screen exited by window {msgCopy.LParam}");
                                    break;
                                }

                            case HSHELL.GETMINRECT:
                                SHELLHOOKINFO minRectInfo = Marshal.PtrToStructure<SHELLHOOKINFO>(msg.LParam);
                                if (Windows.Any(i => i.Handle == minRectInfo.hwnd))
                                {
                                    ApplicationWindow win = Windows.First(wnd => wnd.Handle == minRectInfo.hwnd);
                                    minRectInfo.rc = win.GetButtonRectFromShell();

                                    if (minRectInfo.rc.Width <= 0 && minRectInfo.rc.Height <= 0)
                                    {
                                        break;
                                    }
                                    Marshal.StructureToPtr(minRectInfo, msg.LParam, false);
                                    msg.Result = (IntPtr)1;
                                    ShellLogger.Debug($"TasksService: MinRect {minRectInfo.rc.Width}x{minRectInfo.rc.Height} provided for {win.Handle} ({win.Title})");
                                    return; // return here so the result isnt reset to DefWindowProc
                                }
                                break;

                            // TaskMan needs to return true if we provide our own task manager to prevent explorers.
                            // case HSHELL.TASKMAN:
                            //     SingletonLogger.Instance.Info("TaskMan Message received.");
                            //     break;

                            default:
                                break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    ShellLogger.Error("TasksService: Error in ShellWinProc. ", ex);
                    Debugger.Break();
                }
            }
            else if (msg.Msg == WM_TASKBARCREATEDMESSAGE)
            {
                ShellLogger.Debug("TasksService: TaskbarCreated received, setting ITaskbarList window");
                setTaskbarListHwnd(_HookWin.Handle);
            }
            else if (msg.Msg >= (int)WM.USER)
            {
                // Handle ITaskbarList functions, most not implemented yet

                ApplicationWindow win = null;

                switch (msg.Msg)
                {
                    case (int)WM.USER + 50:
                        // ActivateTab
                        // Also sends WM_SHELLHOOK message
                        ShellLogger.Debug("TasksService: ITaskbarList: ActivateTab HWND:" + msg.LParam);
                        msg.Result = IntPtr.Zero;
                        return;
                    // Deliberately no case for AddTab/DeleteTab (the WM.USER+51/+52 slots the old
                    // ITaskbarList v1 message protocol used for these): confirmed via RetroBar's own
                    // logs that real ITaskbarList3::DeleteTab calls (e.g. from window_manager's/
                    // Tabame's own "skip taskbar" - both just CoCreateInstance(CLSID_TaskbarList) and
                    // call the COM method) never arrive here at all, unlike ActivateTab/
                    // MarkFullscreenWindow below which fire live. Modern AddTab/DeleteTab callers go
                    // through COM/RPC straight into explorer.exe, not the legacy TaskbandHWND-redirected
                    // SendMessage channel this hook intercepts. ApplicationWindow.CanAddToTaskbar's
                    // "ITaskList_Deleted" property check is what real Explorer sets on DeleteTab, but
                    // nothing here can write it - would require overriding the system's CLSID_TaskbarList
                    // COM registration itself, which is out of scope (machine-wide, affects every app).
                    case (int)WM.USER + 60:
                        // MarkFullscreenWindow
                        // Also sends WM_SHELLHOOK message
                        ShellLogger.Debug("TasksService: ITaskbarList: MarkFullscreenWindow HWND:" + msg.LParam + " Entering? " + msg.WParam);
                        FullScreenEventArgs args = new FullScreenEventArgs
                        {
                            Handle = msgCopy.LParam,
                            IsEntering = msg.WParam != IntPtr.Zero
                        };

                        FullScreenChanged?.Invoke(this, args);
                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 64:
                        // SetProgressValue
                        ShellLogger.Debug("TasksService: ITaskbarList: SetProgressValue HWND:" + msg.WParam + " Progress: " + msg.LParam);

                        win = new ApplicationWindow(this, msg.WParam);
                        if (Windows.Contains(win))
                        {
                            win = Windows.First(wnd => wnd.Handle == msgCopy.WParam);
                            win.ProgressValue = (int)msg.LParam;
                        }

                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 65:
                        // SetProgressState
                        ShellLogger.Debug("TasksService: ITaskbarList: SetProgressState HWND:" + msg.WParam + " Flags: " + msg.LParam);

                        win = new ApplicationWindow(this, msg.WParam);
                        if (Windows.Contains(win))
                        {
                            win = Windows.First(wnd => wnd.Handle == msgCopy.WParam);
                            win.ProgressState = (TBPFLAG)msg.LParam;
                        }

                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 67:
                        // RegisterTab - this window is a sub-tab of another; hide it from the taskbar
                        ShellLogger.Debug("TasksService: ITaskbarList: RegisterTab MDI HWND:" + msg.LParam + " Tab HWND: " + msg.WParam);
                        removeWindow(msg.WParam);
                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 68:
                        // UnregisterTab - window is no longer a sub-tab; re-add it if eligible
                        ShellLogger.Debug("TasksService: ITaskbarList: UnregisterTab Tab HWND: " + msg.WParam);
                        addWindow(msg.WParam, ApplicationWindow.WindowState.Inactive, true);
                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 71:
                        // SetTabOrder
                        ShellLogger.Debug("TasksService: ITaskbarList: SetTabOrder HWND:" + msg.WParam + " Before HWND: " + msg.LParam);
                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 72:
                        // SetTabActive
                        ShellLogger.Debug("TasksService: ITaskbarList: SetTabActive HWND:" + msg.WParam);
                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 75:
                        // Unknown
                        ShellLogger.Debug("TasksService: ITaskbarList: Unknown HWND:" + msg.WParam + " LParam: " + msg.LParam);
                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 76:
                        // ThumbBarAddButtons
                        ShellLogger.Debug("TasksService: ITaskbarList: ThumbBarAddButtons HWND:" + msg.WParam);
                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 77:
                        // ThumbBarUpdateButtons
                        ShellLogger.Debug("TasksService: ITaskbarList: ThumbBarUpdateButtons HWND:" + msg.WParam);
                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 78:
                        // ThumbBarSetImageList
                        ShellLogger.Debug("TasksService: ITaskbarList: ThumbBarSetImageList HWND:" + msg.WParam);
                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 79:
                        // SetOverlayIcon - Icon
                        ShellLogger.Debug("TasksService: ITaskbarList: SetOverlayIcon - Icon HWND:" + msg.WParam);

                        win = new ApplicationWindow(this, msg.WParam);
                        if (Windows.Contains(win))
                        {
                            win = Windows.First(wnd => wnd.Handle == msgCopy.WParam);
                            win.SetOverlayIcon(msg.LParam);
                        }

                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 80:
                        // SetThumbnailTooltip
                        ShellLogger.Debug("TasksService: ITaskbarList: SetThumbnailTooltip HWND:" + msg.WParam);
                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 81:
                        // SetThumbnailClip
                        ShellLogger.Debug("TasksService: ITaskbarList: SetThumbnailClip HWND:" + msg.WParam);
                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 85:
                        // SetOverlayIcon - Description
                        ShellLogger.Debug("TasksService: ITaskbarList: SetOverlayIcon - Description HWND:" + msg.WParam);

                        win = new ApplicationWindow(this, msg.WParam);
                        if (Windows.Contains(win))
                        {
                            win = Windows.First(wnd => wnd.Handle == msgCopy.WParam);
                            win.SetOverlayIconDescription(msg.LParam);
                        }

                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 87:
                        // SetTabProperties
                        ShellLogger.Debug("TasksService: ITaskbarList: SetTabProperties HWND:" + msg.WParam);
                        msg.Result = IntPtr.Zero;
                        return;
                    default:
                        ShellLogger.Debug($"TasksService: Unknown ITaskbarList Msg: {msg.Msg} LParam: {msg.LParam} WParam: {msg.WParam}");
                        break;
                }
            }

            handled = false;
        }

        private void MoveEventCallback(IntPtr hWinEventHook, uint eventType, IntPtr hWnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            if (hWnd != IntPtr.Zero && idObject == 0 && idChild == 0)
            {
                if (Windows.Any(i => i.Handle == hWnd))
                {
                    ApplicationWindow win = Windows.First(wnd => wnd.Handle == hWnd);
                    win.SetMonitor();
                }
            }
        }

        private void CloakEventCallback(IntPtr hWinEventHook, uint eventType, IntPtr hWnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            if (hWnd != IntPtr.Zero && idObject == 0 && idChild == 0)
            {
                if (Windows.Any(i => i.Handle == hWnd))
                {
                    ApplicationWindow win = Windows.First(wnd => wnd.Handle == hWnd);
                    ShellLogger.Debug($"TasksService: {(eventType == EVENT_OBJECT_CLOAKED ? "Cloak" : "Uncloak")} event received for {win.Title}");
                    win.SetShowInTaskbar();
                }
                else if (eventType == EVENT_OBJECT_UNCLOAKED)
                {
                    // Window was on another virtual desktop when RetroBar started (or was created
                    // there later) and was never added to Windows. Now that DWM has uncloaked it,
                    // add it so it appears in the taskbar.
                    lock (_windowsLock)
                    {
                        ApplicationWindow win = new ApplicationWindow(this, hWnd);
                        if (win.CanAddToTaskbar && win.ShowInTaskbar)
                        {
                            Windows.Add(win);
                            sendTaskbarButtonCreatedMessage(win.Handle);
                            ShellLogger.Debug($"TasksService: Uncloaked previously-unseen window {hWnd} ({win.Title})");
                        }
                    }

                    // Schedule a deferred EnumWindows scan as a fallback for Win32 apps that
                    // don't generate EVENT_OBJECT_UNCLOAKED reliably. Each uncloak event resets
                    // the timer so we do one scan per desktop-switch batch.
                    ScheduleDeferredWindowScan();
                }
            }
        }

        private void ScheduleDeferredWindowScan()
        {
            if (_desktopSwitchScanTimer == null)
            {
                _desktopSwitchScanTimer = new System.Windows.Threading.DispatcherTimer();
                _desktopSwitchScanTimer.Tick += (s, e) =>
                {
                    _desktopSwitchScanTimer.Stop();
                    ScanForNewWindows();
                };
            }

            _desktopSwitchScanTimer.Stop();
            _desktopSwitchScanTimer.Interval = TimeSpan.FromMilliseconds(500);
            _desktopSwitchScanTimer.Start();
        }

        private void ScanForNewWindows()
        {
            ShellLogger.Debug("TasksService: Deferred scan for previously-unseen windows");
            lock (_windowsLock)
            {
                EnumWindows((hwnd, lParam) =>
                {
                    if (!Windows.Any(i => i.Handle == hwnd))
                    {
                        ApplicationWindow win = new ApplicationWindow(this, hwnd);
                        if (win.CanAddToTaskbar && win.ShowInTaskbar)
                        {
                            Windows.Add(win);
                            sendTaskbarButtonCreatedMessage(win.Handle);
                            ShellLogger.Debug($"TasksService: Deferred scan added previously-unseen window {hwnd} ({win.Title})");
                        }
                    }
                    return true;
                }, 0);
            }
        }

        // Fires for non-virtual desktop switches: UAC prompt, lock screen, fast-user-switch.
        // Does NOT fire for Windows 10/11 virtual desktop switches; those use per-window
        // EVENT_OBJECT_CLOAKED/UNCLOAKED events handled by CloakEventCallback.
        private void DesktopSwitchEventCallback(IntPtr hWinEventHook, uint eventType, IntPtr hWnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            ShellLogger.Debug("TasksService: Desktop switch detected, refreshing window list");
            lock (_windowsLock)
            {
                foreach (ApplicationWindow win in Windows)
                {
                    win.SetShowInTaskbar();
                }
            }
            ScanForNewWindows();
        }

        // set property on hook window that should receive ITaskbarList messages
        private void setTaskbarListHwnd(IntPtr hwndHook)
        {
            bool resetProp = true;

            // get the topmost tray
            IntPtr taskbarHwnd = WindowHelper.FindWindowsTray(IntPtr.Zero);
            
            if (taskbarHwnd == IntPtr.Zero)
            {
                return;
            }

            // if our tray is running, there may also be a second tray running
            IntPtr systemTaskbarHwnd = WindowHelper.FindWindowsTray(taskbarHwnd);

            if (hwndHook == IntPtr.Zero)
            {
                // no target hwnd provided
                // Try to find and use the handle of the Explorer hook window
                resetProp = false;
                hwndHook = getChildHwndByClass(systemTaskbarHwnd == IntPtr.Zero ? taskbarHwnd : systemTaskbarHwnd, "MSTaskSwWClass");
            }

            if (hwndHook == IntPtr.Zero)
            {
                // if still no hwnd to hook, we can't do anything
                return;
            }

            ShellLogger.Debug("TasksService: Adding TaskbandHWND prop to hwnd: " + taskbarHwnd);
            SetProp(taskbarHwnd, "TaskbandHWND", hwndHook);

            // Remove the property from the Explorer taskbar, if it is not the only tray
            if (resetProp && systemTaskbarHwnd != IntPtr.Zero)
            {
                ShellLogger.Debug("TasksService: Removing TaskbandHWND prop from hwnd: " + systemTaskbarHwnd);
                RemoveProp(systemTaskbarHwnd, "TaskbandHWND");
            }
        }

        private IntPtr getChildHwndByClass(IntPtr parentHwnd, string wndClass)
        {
            IntPtr childHwnd = IntPtr.Zero;
            EnumChildWindows(parentHwnd, (hwnd, lParam) =>
            {
                StringBuilder cName = new StringBuilder(256);
                GetClassName(hwnd, cName, cName.Capacity);
                if (cName.ToString() == wndClass)
                {
                    childHwnd = hwnd;
                    return false;
                }

                return true;
            }, 0);

            return childHwnd;
        }

        public ObservableCollection<ApplicationWindow> Windows
        {
            get
            {
                return base.GetValue(windowsProperty) as ObservableCollection<ApplicationWindow>;
            }
            set
            {
                SetValue(windowsProperty, value);
            }
        }

        private DependencyProperty windowsProperty = DependencyProperty.Register("Windows",
            typeof(ObservableCollection<ApplicationWindow>), typeof(TasksService),
            new PropertyMetadata(new ObservableCollection<ApplicationWindow>()));
    }
}
