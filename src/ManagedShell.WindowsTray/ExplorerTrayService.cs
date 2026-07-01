using ManagedShell.Common.Helpers;
using ManagedShell.Common.Logging;
using Microsoft.Win32;
using System;
using System.Runtime.InteropServices;
using static ManagedShell.Interop.NativeMethods;

namespace ManagedShell.WindowsTray
{
    public  class ExplorerTrayService
    {
        private SystrayDelegate trayDelegate;

        public ExplorerTrayService()
        {
        }

        internal void SetSystrayCallback(SystrayDelegate theDelegate)
        {
            trayDelegate = theDelegate;
        }

        internal void Run()
        {
            if (!EnvironmentHelper.IsAppRunningAsShell && trayDelegate != null)
            {
                bool autoTrayEnabled = GetAutoTrayEnabled();
                TrayNotify trayNotify = null;

                // we can't get tray icons that are in the hidden area, so disable that temporarily if enabled.
                // Failure here (e.g. TrayNotify CLSID not registered on this Windows build) must not skip the
                // actual icon read below - it just means hidden icons may not be enumerated this pass.
                if (autoTrayEnabled)
                {
                    try
                    {
                        trayNotify = new TrayNotify();
                        SetAutoTrayEnabled(trayNotify, false);
                    }
                    catch (Exception e)
                    {
                        ShellLogger.Debug($"ExplorerTrayService: Unable to disable auto-tray via ITrayNotify: {e.Message}");
                        trayNotify = null;
                    }
                }

                try
                {
                    GetTrayItems();
                }
                catch (Exception e)
                {
                    ShellLogger.Debug($"ExplorerTrayService: Unable to get items: {e.Message}");
                }

                if (trayNotify != null)
                {
                    try
                    {
                        SetAutoTrayEnabled(trayNotify, true);
                    }
                    catch (Exception e)
                    {
                        ShellLogger.Debug($"ExplorerTrayService: Unable to re-enable auto-tray via ITrayNotify: {e.Message}");
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(trayNotify);
                    }
                }
            }
        }

        private void GetTrayItems()
        {
            IntPtr toolbarHwnd = FindExplorerTrayToolbarHwnd();

            if (toolbarHwnd == IntPtr.Zero)
            {
                ShellLogger.Warning("ExplorerTrayService: Could not find Explorer tray toolbar; trying ITrayNotify callback fallback");
                GetTrayItemsViaCallback();
                return;
            }

            int count = GetNumTrayIcons(toolbarHwnd);

            if (count < 1)
            {
                return;
            }

            GetWindowThreadProcessId(toolbarHwnd, out var processId);
            IntPtr hProcess = OpenProcess(ProcessAccessFlags.All, false, (int)processId);
            IntPtr hBuffer = VirtualAllocEx(hProcess, IntPtr.Zero, (uint)Marshal.SizeOf(new TBBUTTON()), AllocationType.Commit,
                MemoryProtection.ReadWrite);

            for (int i = 0; i < count; i++)
            {
                TrayItem trayItem = GetTrayItem(i, hBuffer, hProcess, toolbarHwnd);

                if (trayItem.hWnd == IntPtr.Zero || !IsWindow(trayItem.hWnd))
                {
                    ShellLogger.Debug($"ExplorerTrayService: Ignored notify icon {trayItem.szIconText} due to invalid handle");
                    continue;
                }

                SafeNotifyIconData nid = GetTrayItemIconData(trayItem);

                if (trayDelegate != null)
                {
                    if (!trayDelegate((uint)NIM.NIM_ADD, nid))
                    {
                        ShellLogger.Debug($"ExplorerTrayService: Ignored notify icon {trayItem.szIconText} hWnd={nid.hWnd} GUID={nid.guidItem}");
                    }
                }
                else
                {
                    ShellLogger.Debug("ExplorerTrayService: trayDelegate is null");
                }
            }

            VirtualFreeEx(hProcess, hBuffer, 0, AllocationType.Release);

            CloseHandle(hProcess);
        }

        private IntPtr FindExplorerTrayToolbarHwnd()
        {
            IntPtr hwnd = FindWindow("Shell_TrayWnd", "");

            if (hwnd == IntPtr.Zero)
            {
                ShellLogger.Debug("ExplorerTrayService: Shell_TrayWnd not found");
                return IntPtr.Zero;
            }

            hwnd = FindWindowEx(hwnd, IntPtr.Zero, "TrayNotifyWnd", "");

            if (hwnd == IntPtr.Zero)
            {
                ShellLogger.Debug("ExplorerTrayService: TrayNotifyWnd not found under Shell_TrayWnd");
                return IntPtr.Zero;
            }

            IntPtr sysPager = FindWindowEx(hwnd, IntPtr.Zero, "SysPager", "");

            if (sysPager == IntPtr.Zero)
            {
                ShellLogger.Debug("ExplorerTrayService: SysPager not found under TrayNotifyWnd — trying ToolbarWindow32 directly");
                hwnd = FindWindowEx(hwnd, IntPtr.Zero, "ToolbarWindow32", IntPtr.Zero);
                if (hwnd == IntPtr.Zero)
                    ShellLogger.Debug("ExplorerTrayService: ToolbarWindow32 not found directly under TrayNotifyWnd either");
                return hwnd;
            }

            hwnd = FindWindowEx(sysPager, IntPtr.Zero, "ToolbarWindow32", IntPtr.Zero);

            if (hwnd == IntPtr.Zero)
                ShellLogger.Debug("ExplorerTrayService: ToolbarWindow32 not found under SysPager");

            return hwnd;
        }

        private int GetNumTrayIcons(IntPtr toolbarHwnd)
        {
            return (int)SendMessage(toolbarHwnd, (int)TB.BUTTONCOUNT, IntPtr.Zero, IntPtr.Zero);
        }

        private TrayItem GetTrayItem(int i, IntPtr hBuffer, IntPtr hProcess, IntPtr toolbarHwnd)
        {
            TBBUTTON tbButton = new TBBUTTON();
            TrayItem trayItem = new TrayItem();
            IntPtr hTBButton = Marshal.AllocHGlobal(Marshal.SizeOf(tbButton));
            IntPtr hTrayItem = Marshal.AllocHGlobal(Marshal.SizeOf(trayItem));

            IntPtr msgSuccess = SendMessage(toolbarHwnd, (int)TB.GETBUTTON, (IntPtr)i, hBuffer);
            if (ReadProcessMemory(hProcess, hBuffer, hTBButton, Marshal.SizeOf(tbButton), out _))
            {
                tbButton = (TBBUTTON)Marshal.PtrToStructure(hTBButton, typeof(TBBUTTON));

                if (tbButton.dwData != UIntPtr.Zero)
                {
                    if (ReadProcessMemory(hProcess, tbButton.dwData, hTrayItem, Marshal.SizeOf(trayItem), out _))
                    {
                        trayItem = (TrayItem)Marshal.PtrToStructure(hTrayItem, typeof(TrayItem));

                        if ((tbButton.fsState & TBSTATE_HIDDEN) != 0)
                        {
                            trayItem.dwState = 1;
                        }
                        else
                        {
                            trayItem.dwState = 0;
                        }

                        ShellLogger.Debug(
                            $"ExplorerTrayService: Got tray item: {trayItem.szIconText}");
                    }
                }
            }

            return trayItem;
        }

        private SafeNotifyIconData GetTrayItemIconData(TrayItem trayItem)
        {
            SafeNotifyIconData nid = new SafeNotifyIconData();

            nid.hWnd = trayItem.hWnd;
            nid.uID = trayItem.uID;
            nid.uCallbackMessage = trayItem.uCallbackMessage;
            nid.szTip = trayItem.szIconText;
            nid.hIcon = trayItem.hIcon;
            nid.uVersion = trayItem.uVersion;
            nid.guidItem = trayItem.guidItem;
            nid.dwState = (int)trayItem.dwState;
            nid.uFlags = NIF.GUID | NIF.MESSAGE | NIF.TIP | NIF.STATE;

            if (nid.hIcon != IntPtr.Zero)
            {
                nid.uFlags |= NIF.ICON;
            }
            else
            {
                ShellLogger.Warning($"ExplorerTrayService: Unable to use {trayItem.szIconText} icon handle for NOTIFYICONDATA struct");
            }

            return nid;
        }

        private void GetTrayItemsViaCallback()
        {
            try
            {
                TrayNotify trayNotify = new TrayNotify();
                var cb = new NotificationCB(trayDelegate);

                if (EnvironmentHelper.IsWindows8OrBetter)
                {
                    var iface = (ITrayNotify)trayNotify;
                    iface.RegisterCallback(cb, out ulong handle);
                    iface.UnregisterCallback(handle);
                }
                else
                {
                    var iface = (ITrayNotifyLegacy)trayNotify;
                    iface.RegisterCallback(cb);
                }

                ShellLogger.Info($"ExplorerTrayService: ITrayNotify callback pre-populated {cb.Count} icon(s)");
                Marshal.ReleaseComObject(trayNotify);
            }
            catch (Exception e)
            {
                ShellLogger.Warning($"ExplorerTrayService: ITrayNotify callback fallback failed: {e.Message}; existing icons will not be pre-populated");
            }
        }

        private class NotificationCB : INotificationCB
        {
            private readonly SystrayDelegate _trayDelegate;
            public int Count { get; private set; }

            public NotificationCB(SystrayDelegate trayDelegate)
            {
                _trayDelegate = trayDelegate;
            }

            public void Notify(uint nEvent, ref NOTIFYITEM item)
            {
                if (item.hWnd == IntPtr.Zero) return;

                var nid = new SafeNotifyIconData
                {
                    hWnd = item.hWnd,
                    uID = item.uID,
                    guidItem = item.guidItem,
                    hIcon = item.hIcon,
                    szTip = item.pszIconText,
                    uFlags = NIF.TIP | NIF.MESSAGE | NIF.STATE
                };

                if (item.hIcon != IntPtr.Zero)
                    nid.uFlags |= NIF.ICON;

                if (item.guidItem != Guid.Empty)
                    nid.uFlags |= NIF.GUID;

                ShellLogger.Debug($"ExplorerTrayService: ITrayNotify nEvent={nEvent} icon={item.pszIconText} exe={item.pszExeName} hWnd={item.hWnd} GUID={item.guidItem}");

                if (_trayDelegate != null && _trayDelegate((uint)NIM.NIM_ADD, nid))
                    Count++;
            }
        }

        private bool GetAutoTrayEnabled()
        {
            int enableAutoTray = 1;

            try
            {
                RegistryKey explorerKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer", false);

                if (explorerKey != null)
                {
                    var enableAutoTrayValue = explorerKey.GetValue("EnableAutoTray");

                    if (enableAutoTrayValue != null)
                    {
                        enableAutoTray = Convert.ToInt32(enableAutoTrayValue);
                    }
                }
            }
            catch (Exception e)
            {
                ShellLogger.Debug($"ExplorerTrayService: Unable to get EnableAutoTray setting: {e.Message}");
            }

            return enableAutoTray == 1;
        }

        private void SetAutoTrayEnabled(TrayNotify trayNotify, bool enabled)
        {
            try
            {
                if (EnvironmentHelper.IsWindows8OrBetter)
                {
                    var trayNotifyInstance = (ITrayNotify)trayNotify;
                    trayNotifyInstance.EnableAutoTray(enabled);
                }
                else
                {
                    var trayNotifyInstance = (ITrayNotifyLegacy)trayNotify;
                    trayNotifyInstance.EnableAutoTray(enabled);
                }
            }
            catch (Exception e)
            {
                ShellLogger.Debug($"ExplorerTrayService: Unable to set EnableAutoTray setting: {e.Message}");
            }
        }

        private const byte TBSTATE_HIDDEN = 8;

        private enum TB : uint
        {
            GETBUTTON = WM.USER + 23,
            BUTTONCOUNT = WM.USER + 24
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct TrayItem
        {
            public IntPtr hWnd;
            public uint uID;
            public uint uCallbackMessage;
            public uint dwState;
            public uint uVersion;
            public IntPtr hIcon;
            public IntPtr uIconDemoteTimerID;
            public uint dwUserPref;
            public uint dwLastSoundTime;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szExeName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szIconText;
            public uint uNumSeconds;
            public Guid guidItem;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct TBBUTTON
        {
            public int iBitmap;
            public int idCommand;
            [StructLayout(LayoutKind.Explicit)]
            private struct TBBUTTON_U
            {
                [FieldOffset(0)] public byte fsState;
                [FieldOffset(1)] public byte fsStyle;
                [FieldOffset(0)] private IntPtr bReserved;
            }
            private TBBUTTON_U union;
            public byte fsState { get { return union.fsState; } set { union.fsState = value; } }
            public byte fsStyle { get { return union.fsStyle; } set { union.fsStyle = value; } }
            public UIntPtr dwData;
            public IntPtr iString;
        }
    }
}
