using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Windows.Automation;
using System.IO;

namespace GhostBusterPlus
{
    class Program
    {
        // Windows API constants and structures
        private const int WH_MOUSE_LL = 14;
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_MOUSEWHEEL = 0x020A;
        private const int WM_MOUSEHWHEEL = 0x020E;
        private const int WM_LBUTTONDOWN = 0x0201;
        private const int WM_LBUTTONUP = 0x0202;
        private const int WM_RBUTTONDOWN = 0x0204;
        private const int WM_RBUTTONUP = 0x0205;
        private const int WM_MBUTTONDOWN = 0x0207;
        private const int WM_MBUTTONUP = 0x0208;
        private const int WM_XBUTTONDOWN = 0x020B;
        private const int WM_XBUTTONUP = 0x020C;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_HOTKEY = 0x0312;
        private const int HOTKEY_ID_TOGGLE = 1; // For Ctrl+Shift+Z
        private const int HOTKEY_ID_RESTART = 2; // For Ctrl+Shift+X
        private const int HOTKEY_ID_REFRESH = 3; // For Ctrl+Shift+D

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MSLLHOOKSTRUCT
        {
            public POINT pt;
            public int mouseData;
            public int flags;
            public int time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        // Windows API functions
        [DllImport("user32.dll")]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll")]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll")]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        private static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        [DllImport("kernel32.dll")]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("kernel32.dll")]
        private static extern void Beep(int frequency, int duration);

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);
        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        // Constants for mouse events
        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;

        // Global variables
        private static IntPtr mouseHook = IntPtr.Zero;
        private static IntPtr keyboardHook = IntPtr.Zero;
        private static LowLevelMouseProc mouseProc;
        private static LowLevelKeyboardProc keyboardProc;
        private static NotifyIcon trayIcon;
        private static ContextMenuStrip trayMenu;
        private static ToolStripMenuItem toggleMonitoringMenuItem;
        private static System.Windows.Forms.Timer eventTimer;
        private static HotkeyWindow hotkeyWindow;
        private static IntPtr lastWindowId = IntPtr.Zero;
        private static long lastEventTime = 0;
        private static bool hasEvent = false;
        private static bool isMonitoring = true;
        private static bool isMouseButtonHeld = false;
        private static bool isRestoringFocus = false;
        private static int eventTimeout = 1500; // Default to 1500 ms, will be loaded from file
        private static int targetX = 0; // Will be set dynamically
        private static int targetY = 0; // Will be set dynamically
        private static int originalX = 0;
        private static int originalY = 0;
        private static IntPtr originalWindowId = IntPtr.Zero;
        private static int lastMouseX = 0;
        private static int lastMouseY = 0;

        // UI Automation variables for scroll detection
        private static AutomationElement currentScrollElement = null;
        private static AutomationPropertyChangedEventHandler scrollEventHandler = null;
        private static bool usingMouseWheelFallback = false;

        // Variables for window move/resize detection
        private static RECT lastWindowPosition = new RECT();
        private static bool isWindowMovingOrResizing = false;
        private static bool wasWindowMovedOrResized = false;

        // Timeout options and persistence
        private static readonly int[] timeoutOptions = { 1000, 1500, 2000, 3000, 5000, 10000, 20000 };
        private static ToolStripMenuItem[] timeoutMenuItems;
        private static readonly string timeoutFilePath = "timeout.txt";

        // Exception for scroll detection timeout
        private class ScrollDetectionTimeoutException : Exception
        {
            public ScrollDetectionTimeoutException(string message) : base(message) { }
        }

        // Message-only window for hotkey processing
        private class HotkeyWindow : NativeWindow
        {
            public HotkeyWindow()
            {
                CreateHandle(new CreateParams());
            }

            protected override void WndProc(ref Message m)
            {
                if (m.Msg == WM_HOTKEY)
                {
                    if (m.WParam == (IntPtr)HOTKEY_ID_TOGGLE)
                    {
                        Debug.WriteLine("WM_HOTKEY message received: Ctrl+Shift+Z");
                        ToggleMonitoring(null, null);
                    }
                    else if (m.WParam == (IntPtr)HOTKEY_ID_RESTART)
                    {
                        Debug.WriteLine("WM_HOTKEY message received: Ctrl+Shift+X");
                        RestartEInkPlus(null, null);
                    }
                    else if (m.WParam == (IntPtr)HOTKEY_ID_REFRESH)
                    {
                        Debug.WriteLine("WM_HOTKEY message received: Ctrl+Shift+D");
                        PerformClickAndRestore(null, null);
                    }
                }
                base.WndProc(ref m);
            }
        }

        static void Main()
        {
            Debug.WriteLine("Starting GhostBusterPlus application.");

            // Load the saved timeout value
            LoadTimeoutValue();

            // Initialize system tray
            trayMenu = new ContextMenuStrip();
            toggleMonitoringMenuItem = new ToolStripMenuItem("Disable auto-GhostBust (Ctrl+Shift+Z)", null, ToggleMonitoring);
            trayMenu.Items.Add(toggleMonitoringMenuItem);
            trayMenu.Items.Add("Restart EInkPlus (Ctrl+Shift+X)", null, RestartEInkPlus);
            trayMenu.Items.Add("Refresh (Ctrl+Shift+D)", null, PerformClickAndRestore);

            // Add Timeout sub-menu
            var timeoutMenu = new ToolStripMenuItem("Set Timeout");
            timeoutMenuItems = new ToolStripMenuItem[timeoutOptions.Length];
            for (int i = 0; i < timeoutOptions.Length; i++)
            {
                int timeoutValue = timeoutOptions[i];
                timeoutMenuItems[i] = new ToolStripMenuItem($"{timeoutValue} ms", null, (s, e) => SetTimeout(timeoutValue));
                if (timeoutValue == eventTimeout)
                {
                    timeoutMenuItems[i].Font = new Font(timeoutMenuItems[i].Font, FontStyle.Bold);
                }
                timeoutMenu.DropDownItems.Add(timeoutMenuItems[i]);
            }
            trayMenu.Items.Add(timeoutMenu);

            trayMenu.Items.Add(new ToolStripSeparator());
            trayMenu.Items.Add("About...", null, ShowAboutDialog);
            trayMenu.Items.Add("Exit", null, ExitApplication);

            trayIcon = new NotifyIcon
            {
                Icon = LoadCustomIcon(),
                ContextMenuStrip = trayMenu,
                Visible = true,
                Text = "Click and Restore Script"
            };
            trayIcon.DoubleClick += ToggleMonitoring;
            Debug.WriteLine("System tray initialized.");

            // Initialize hotkey window and register global hotkeys
            hotkeyWindow = new HotkeyWindow();
            RegisterGlobalHotKey();

            // Install mouse and keyboard hooks
            mouseProc = MouseHookProc;
            keyboardProc = KeyboardHookProc;
            mouseHook = SetWindowsHookEx(WH_MOUSE_LL, mouseProc, GetModuleHandle(null), 0);
            keyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, keyboardProc, GetModuleHandle(null), 0);

            if (mouseHook == IntPtr.Zero || keyboardHook == IntPtr.Zero)
            {
                Debug.WriteLine("Failed to set mouse or keyboard hooks.");
                MessageBox.Show("Failed to set hooks!");
                Application.Exit();
                return;
            }
            Debug.WriteLine("Mouse and keyboard hooks installed successfully.");

            // Initialize event timer (250ms interval)
            eventTimer = new System.Windows.Forms.Timer { Interval = 250 };
            eventTimer.Tick += CheckEventTimeout;
            eventTimer.Tick += CheckWindowChange;
            eventTimer.Tick += CheckMouseMovement;
            eventTimer.Tick += UpdateScrollDetection;
            eventTimer.Tick += CheckWindowMoveOrResize;
            eventTimer.Start();
            Debug.WriteLine("Event timer started with 250ms interval.");

            Application.Run();
            Debug.WriteLine("Application exited.");
        }

        private static void LoadTimeoutValue()
        {
            try
            {
                if (File.Exists(timeoutFilePath))
                {
                    string timeoutStr = File.ReadAllText(timeoutFilePath).Trim();
                    if (int.TryParse(timeoutStr, out int savedTimeout))
                    {
                        // Ensure the saved timeout is one of the valid options
                        if (Array.IndexOf(timeoutOptions, savedTimeout) != -1)
                        {
                            eventTimeout = savedTimeout;
                            Debug.WriteLine($"Loaded timeout value: {eventTimeout} ms");
                            return;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading timeout value: {ex.Message}");
            }

            // Default to 1500 ms if file doesn't exist or is invalid
            eventTimeout = 1500;
            Debug.WriteLine($"Using default timeout value: {eventTimeout} ms");
        }

        private static void SaveTimeoutValue()
        {
            try
            {
                File.WriteAllText(timeoutFilePath, eventTimeout.ToString());
                Debug.WriteLine($"Saved timeout value: {eventTimeout} ms");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error saving timeout value: {ex.Message}");
            }
        }

        private static void SetTimeout(int newTimeout)
        {
            eventTimeout = newTimeout;
            SaveTimeoutValue();

            // Update the menu to bold the selected timeout
            foreach (var menuItem in timeoutMenuItems)
            {
                int menuTimeout = int.Parse(menuItem.Text.Replace(" ms", ""));
                if (menuTimeout == eventTimeout)
                {
                    menuItem.Font = new Font(menuItem.Font, FontStyle.Bold);
                }
                else
                {
                    menuItem.Font = new Font(menuItem.Font, FontStyle.Regular);
                }
            }

            Debug.WriteLine($"Timeout set to: {eventTimeout} ms");
        }

        private static Icon LoadCustomIcon()
        {
            try
            {
                return new Icon("icon.ico");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to load custom icon: {ex.Message}. Using default icon.");
                return SystemIcons.Application;
            }
        }

        private static void RegisterGlobalHotKey()
        {
            // Register Ctrl+Shift+Z (0x5A = Z key, MOD_CONTROL | MOD_SHIFT = 0x0001 | 0x0002 = 0x0006)
            bool successToggle = RegisterHotKey(hotkeyWindow.Handle, HOTKEY_ID_TOGGLE, 0x0006, 0x5A);
            Debug.WriteLine($"Hotkey registration (Ctrl+Shift+Z): {(successToggle ? "Success" : "Failed")}");
            if (!successToggle)
            {
                MessageBox.Show("Failed to register hotkey Ctrl+Shift+Z!");
            }

            // Register Ctrl+Shift+X (0x58 = X key, MOD_CONTROL | MOD_SHIFT = 0x0001 | 0x0002 = 0x0006)
            bool successRestart = RegisterHotKey(hotkeyWindow.Handle, HOTKEY_ID_RESTART, 0x0006, 0x58);
            Debug.WriteLine($"Hotkey registration (Ctrl+Shift+X): {(successRestart ? "Success" : "Failed")}");
            if (!successRestart)
            {
                MessageBox.Show("Failed to register hotkey Ctrl+Shift+X!");
            }

            // Register Ctrl+Shift+D (0x44 = D key, MOD_CONTROL | MOD_SHIFT = 0x0001 | 0x0002 = 0x0006)
            bool successRefresh = RegisterHotKey(hotkeyWindow.Handle, HOTKEY_ID_REFRESH, 0x0006, 0x44);
            Debug.WriteLine($"Hotkey registration (Ctrl+Shift+D): {(successRefresh ? "Success" : "Failed")}");
            if (!successRefresh)
            {
                MessageBox.Show("Failed to register hotkey Ctrl+Shift+D!");
            }
        }

        private static void UnregisterGlobalHotKey()
        {
            UnregisterHotKey(hotkeyWindow.Handle, HOTKEY_ID_TOGGLE);
            UnregisterHotKey(hotkeyWindow.Handle, HOTKEY_ID_RESTART);
            UnregisterHotKey(hotkeyWindow.Handle, HOTKEY_ID_REFRESH);
            Debug.WriteLine("Unregistered global hotkeys Ctrl+Shift+Z, Ctrl+Shift+X, and Ctrl+Shift+D.");
        }

        private static IntPtr KeyboardHookProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                // Delay event if any key is pressed
                if (hasEvent)
                {
                    lastEventTime = Stopwatch.GetTimestamp();
                    Debug.WriteLine("Keyboard activity detected, delaying event. Last event time updated.");
                }
            }
            return CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam);
        }

        private static IntPtr MouseHookProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                int msg = wParam.ToInt32();
                MSLLHOOKSTRUCT mouseStruct = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam);

                // Detect mouse wheel events (fallback mechanism)
                if (usingMouseWheelFallback && (msg == WM_MOUSEWHEEL || msg == WM_MOUSEHWHEEL))
                {
                    lastEventTime = Stopwatch.GetTimestamp();
                    hasEvent = true;
                }

                // Detect mouse clicks and held state
                if (msg == WM_LBUTTONDOWN || msg == WM_RBUTTONDOWN || msg == WM_MBUTTONDOWN ||
                    msg == WM_XBUTTONDOWN)
                {
                    if (hasEvent)
                        lastEventTime = Stopwatch.GetTimestamp();
                    isMouseButtonHeld = true;
                    Debug.WriteLine("Mouse button down detected, delaying event if active. Last event time updated.");
                }
                if (msg == WM_LBUTTONUP || msg == WM_RBUTTONUP || msg == WM_MBUTTONUP ||
                    msg == WM_XBUTTONUP)
                {
                    isMouseButtonHeld = false;
                    Debug.WriteLine("Mouse button up detected.");
                }
            }
            return CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam);
        }

        private static void CheckMouseMovement(object sender, EventArgs e)
        {
            GetCursorPos(out POINT currentPos);
            if (lastMouseX == 0 && lastMouseY == 0)
            {
                lastMouseX = currentPos.X;
                lastMouseY = currentPos.Y;
                Debug.WriteLine($"Initialized mouse position: ({lastMouseX}, {lastMouseY})");
                return;
            }
            if (currentPos.X != lastMouseX || currentPos.Y != lastMouseY)
            {
                if (hasEvent)
                {
                    lastEventTime = Stopwatch.GetTimestamp();
                }
                lastMouseX = currentPos.X;
                lastMouseY = currentPos.Y;
            }
        }

        private static void CheckWindowChange(object sender, EventArgs e)
        {
            if (!isMonitoring)
                return;

            // Suppress focus change events caused by the program itself
            if (isRestoringFocus)
            {
                Debug.WriteLine("Suppressing window change event due to program-initiated focus restoration.");
                return;
            }

            IntPtr currentWindowId = GetForegroundWindow();
            string windowClass = GetWindowClassName(currentWindowId);
            if (windowClass == "Windows.UI.Core.CoreWindow" || windowClass == "Shell_TrayWnd" ||
                windowClass == "SysListView32" || windowClass == "ClockFlyoutWindow" ||
                windowClass == "TrayClockWClass" || windowClass == "TrayShowDesktopButtonWClass" ||
                windowClass == "NotifyIconOverflowWindow" ||
                windowClass.Contains("HwndWrapper[LenovoGen4.FlyoutButton"))
            {
                Debug.WriteLine($"Ignored window change to system UI or FlyoutButton window: {windowClass}");
                return;
            }

            if (currentWindowId != lastWindowId && lastWindowId != IntPtr.Zero)
            {
                lastEventTime = Stopwatch.GetTimestamp();
                hasEvent = true;
                Debug.WriteLine($"Window change detected from {lastWindowId} to {currentWindowId} (class: {windowClass}). Event triggered.");
            }
            lastWindowId = currentWindowId;
        }

        private static void CheckWindowMoveOrResize(object sender, EventArgs e)
        {
            if (!isMonitoring)
            {
                isWindowMovingOrResizing = false;
                wasWindowMovedOrResized = false;
                lastWindowPosition = new RECT();
                return;
            }

            // Get the current foreground window
            IntPtr foregroundWindow = GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero)
            {
                Debug.WriteLine("No foreground window found for move/resize detection.");
                return;
            }

            // Skip if the window is one we should ignore (e.g., system UI windows)
            string windowClass = GetWindowClassName(foregroundWindow);
            if (windowClass == "Windows.UI.Core.CoreWindow" || windowClass == "Shell_TrayWnd" ||
                windowClass == "SysListView32" || windowClass == "ClockFlyoutWindow" ||
                windowClass == "TrayClockWClass" || windowClass == "TrayShowDesktopButtonWClass" ||
                windowClass == "NotifyIconOverflowWindow" ||
                windowClass.Contains("HwndWrapper[LenovoGen4.FlyoutButton"))
            {
                return;
            }

            try
            {
                AutomationElement windowElement = AutomationElement.FromHandle(foregroundWindow);
                if (windowElement != null)
                {
                    var currentPosition = windowElement.Current.BoundingRectangle;
                    RECT currentRect = new RECT
                    {
                        Left = (int)currentPosition.Left,
                        Top = (int)currentPosition.Top,
                        Right = (int)currentPosition.Right,
                        Bottom = (int)currentPosition.Bottom
                    };

                    // Check if the window position or size has changed since the last tick
                    bool positionOrSizeChanged = currentRect.Left != lastWindowPosition.Left ||
                                                 currentRect.Top != lastWindowPosition.Top ||
                                                 currentRect.Right != lastWindowPosition.Right ||
                                                 currentRect.Bottom != lastWindowPosition.Bottom;

                    if (positionOrSizeChanged)
                    {
                        // Window is moving or resizing
                        isWindowMovingOrResizing = true;
                        wasWindowMovedOrResized = true;
                        if (hasEvent)
                        {
                            lastEventTime = Stopwatch.GetTimestamp();
                            Debug.WriteLine("Window move or resize detected, delaying event. Last event time updated.");
                        }
                    }
                    else if (isWindowMovingOrResizing)
                    {
                        // Moving or resizing has stopped
                        isWindowMovingOrResizing = false;
                        if (wasWindowMovedOrResized)
                        {
                            lastEventTime = Stopwatch.GetTimestamp();
                            hasEvent = true;
                            Debug.WriteLine("Window move or resize stopped. Event triggered.");
                            wasWindowMovedOrResized = false;
                        }
                    }

                    // Update the last known position
                    lastWindowPosition = currentRect;
                }
                else
                {
                    Debug.WriteLine($"Failed to get AutomationElement for window: {foregroundWindow} during move/resize detection.");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error checking window move/resize: {ex.Message}");
            }
        }

        private static void UpdateScrollDetection(object sender, EventArgs e)
        {
            if (!isMonitoring)
            {
                // Remove any existing scroll event handler when not monitoring
                if (scrollEventHandler != null && currentScrollElement != null)
                {
                    try
                    {
                        Automation.RemoveAutomationPropertyChangedEventHandler(currentScrollElement, scrollEventHandler);
                        Debug.WriteLine("Removed scroll event handler due to monitoring disabled.");
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error removing scroll event handler: {ex.Message}");
                    }
                    scrollEventHandler = null;
                    currentScrollElement = null;
                }
                usingMouseWheelFallback = false;
                return;
            }

            // Get the current foreground window
            IntPtr foregroundWindow = GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero)
            {
                Debug.WriteLine("No foreground window found for scroll detection.");
                return;
            }

            // Check if the foreground window has changed or if we haven't set up scroll detection yet
            if (foregroundWindow != lastWindowId || (currentScrollElement == null && !usingMouseWheelFallback))
            {
                // Remove any existing scroll event handler
                if (scrollEventHandler != null && currentScrollElement != null)
                {
                    try
                    {
                        Automation.RemoveAutomationPropertyChangedEventHandler(currentScrollElement, scrollEventHandler);
                        Debug.WriteLine("Removed previous scroll event handler.");
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error removing previous scroll event handler: {ex.Message}");
                    }
                }
                currentScrollElement = null;
                scrollEventHandler = null;
                usingMouseWheelFallback = false;

                // Try to set up scroll detection on the new foreground window
                try
                {
                    AutomationElement windowElement = AutomationElement.FromHandle(foregroundWindow);
                    if (windowElement != null)
                    {
                        // Get the current cursor position to find the element under it
                        GetCursorPos(out POINT cursorPos);
                        Debug.WriteLine($"Cursor position for scroll detection: ({cursorPos.X}, {cursorPos.Y})");

                        // Find the element under the cursor
                        AutomationElement elementUnderCursor = AutomationElement.RootElement.FindFirst(TreeScope.Children,
                            new PropertyCondition(AutomationElement.NativeWindowHandleProperty, (int)foregroundWindow));
                        if (elementUnderCursor != null)
                        {
                            // Traverse the tree to find the deepest scrollable control under the cursor
                            AutomationElement scrollableElement;
                            try
                            {
                                scrollableElement = FindScrollableControlUnderCursor(elementUnderCursor, cursorPos.X, cursorPos.Y);
                            }
                            catch (ScrollDetectionTimeoutException ex)
                            {
                                Debug.WriteLine($"Scroll detection timed out after 200ms: {ex.Message}. Falling back to mouse wheel detection.");
                                currentScrollElement = null;
                                scrollEventHandler = null;
                                usingMouseWheelFallback = true;
                                return;
                            }

                            if (scrollableElement != null)
                            {
                                currentScrollElement = scrollableElement;
                                scrollEventHandler = new AutomationPropertyChangedEventHandler(OnScrollPropertyChanged);
                                Automation.AddAutomationPropertyChangedEventHandler(scrollableElement, TreeScope.Element,
                                    scrollEventHandler, ScrollPattern.VerticalScrollPercentProperty, ScrollPattern.HorizontalScrollPercentProperty);
                                Debug.WriteLine($"Set up scroll detection on window: {foregroundWindow}, class: {GetWindowClassName(foregroundWindow)}, control: {scrollableElement.Current.ControlType.ProgrammaticName}");
                                usingMouseWheelFallback = false;
                            }
                            else
                            {
                                Debug.WriteLine($"No scrollable control found under cursor in window: {foregroundWindow}, class: {GetWindowClassName(foregroundWindow)}. Falling back to mouse wheel detection.");
                                currentScrollElement = null;
                                scrollEventHandler = null;
                                usingMouseWheelFallback = true;
                            }
                        }
                        else
                        {
                            Debug.WriteLine($"Failed to find AutomationElement for window: {foregroundWindow}. Falling back to mouse wheel detection.");
                            currentScrollElement = null;
                            scrollEventHandler = null;
                            usingMouseWheelFallback = true;
                        }
                    }
                    else
                    {
                        Debug.WriteLine($"Failed to get AutomationElement for window: {foregroundWindow}. Falling back to mouse wheel detection.");
                        currentScrollElement = null;
                        scrollEventHandler = null;
                        usingMouseWheelFallback = true;
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error setting up scroll detection: {ex.Message}. Falling back to mouse wheel detection.");
                    currentScrollElement = null;
                    scrollEventHandler = null;
                    usingMouseWheelFallback = true;
                }
            }
        }

        private static AutomationElement FindScrollableControlUnderCursor(AutomationElement element, int cursorX, int cursorY)
        {
            if (element == null)
                return null;

            var stopwatch = Stopwatch.StartNew();

            try
            {
                // Check if the current element is scrollable
                if (stopwatch.ElapsedMilliseconds > 200)
                {
                    throw new ScrollDetectionTimeoutException("Traversal exceeded 200ms timeout.");
                }

                if (element.TryGetCurrentPattern(ScrollPattern.Pattern, out object patternObj))
                {
                    ScrollPattern scrollPattern = patternObj as ScrollPattern;
                    if (scrollPattern != null && (scrollPattern.Current.VerticallyScrollable || scrollPattern.Current.HorizontallyScrollable))
                    {
                        // Check if the cursor is within the element's bounding rectangle
                        var rect = element.Current.BoundingRectangle;
                        if (cursorX >= rect.Left && cursorX <= rect.Right && cursorY >= rect.Top && cursorY <= rect.Bottom)
                        {
                            Debug.WriteLine($"Found scrollable control: {element.Current.ControlType.ProgrammaticName}, BoundingRectangle: {rect}");
                            return element;
                        }
                    }
                }

                // Recursively search children
                AutomationElementCollection children = element.FindAll(TreeScope.Children, Condition.TrueCondition);
                foreach (AutomationElement child in children)
                {
                    if (stopwatch.ElapsedMilliseconds > 200)
                    {
                        throw new ScrollDetectionTimeoutException("Traversal exceeded 200ms timeout.");
                    }

                    AutomationElement scrollableChild = FindScrollableControlUnderCursor(child, cursorX, cursorY);
                    if (scrollableChild != null)
                        return scrollableChild;
                }
            }
            catch (Exception ex)
            {
                if (ex is ScrollDetectionTimeoutException)
                    throw; // Re-throw timeout exception to handle in UpdateScrollDetection
                Debug.WriteLine($"Error traversing UI Automation tree for scrollable control: {ex.Message}");
            }
            finally
            {
                stopwatch.Stop();
            }

            return null;
        }

        private static void OnScrollPropertyChanged(object sender, AutomationPropertyChangedEventArgs e)
        {
            if (e.Property == ScrollPattern.VerticalScrollPercentProperty || e.Property == ScrollPattern.HorizontalScrollPercentProperty)
            {
                lastEventTime = Stopwatch.GetTimestamp();
                hasEvent = true;
                Debug.WriteLine($"Scroll detected in window (Property: {e.Property.ProgrammaticName}, NewValue: {e.NewValue}). Last event time updated.");
            }
        }

        private static string GetWindowClassName(IntPtr hWnd)
        {
            var sb = new StringBuilder(256);
            GetClassName(hWnd, sb, 256);
            return sb.ToString();
        }

        private static void CheckEventTimeout(object sender, EventArgs e)
        {
            if (!isMonitoring)
                return;

            if (hasEvent)
            {
                if (isMouseButtonHeld)
                {
                    lastEventTime = Stopwatch.GetTimestamp();
                    Debug.WriteLine("Mouse button held, delaying event timeout.");
                    return;
                }
                long elapsed = (Stopwatch.GetTimestamp() - lastEventTime) * 1000 / Stopwatch.Frequency;
                if (elapsed >= eventTimeout)
                {
                    Debug.WriteLine($"Event timeout reached ({elapsed}ms elapsed). Performing click and restore.");
                    PerformClickAndRestore(null, null);
                    hasEvent = false;
                }
            }
        }

        private static bool IsButtonAvailable()
        {
            int tempTargetX = 0;
            int tempTargetY = 0;

            // Find the FlyoutButton window
            IntPtr flyoutButtonHandle = IntPtr.Zero;
            EnumWindows((hWnd, lParam) =>
            {
                var sb = new StringBuilder(256);
                GetClassName(hWnd, sb, 256);
                string className = sb.ToString();
                if (className.Contains("LenovoGen4.FlyoutButton"))
                {
                    flyoutButtonHandle = hWnd;
                    return false; // Stop enumeration
                }
                return true; // Continue enumeration
            }, IntPtr.Zero);

            if (flyoutButtonHandle != IntPtr.Zero)
            {
                Debug.WriteLine($"Found FlyoutButton window: HWND={flyoutButtonHandle}");

                // Use UI Automation to find the button and get its ClickablePoint
                try
                {
                    AutomationElement flyoutWindow = AutomationElement.FromHandle(flyoutButtonHandle);
                    if (flyoutWindow != null)
                    {
                        Debug.WriteLine("Found FlyoutButton window via UI Automation.");

                        // Find the button control
                        Condition buttonCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button);
                        AutomationElement button = flyoutWindow.FindFirst(TreeScope.Descendants, buttonCondition);

                        if (button != null)
                        {
                            Debug.WriteLine($"Found button control: Name='{button.Current.Name}', ControlType='{button.Current.ControlType.ProgrammaticName}', BoundingRectangle='{button.Current.BoundingRectangle}'");

                            // Get the ClickablePoint
                            var clickablePoint = button.GetClickablePoint();
                            tempTargetX = (int)clickablePoint.X;
                            tempTargetY = (int)clickablePoint.Y;
                            Debug.WriteLine($"Computed ClickablePoint: ({tempTargetX}, {tempTargetY})");
                        }
                        else
                        {
                            Debug.WriteLine("Button control not found via UI Automation.");
                        }
                    }
                    else
                    {
                        Debug.WriteLine("Failed to initialize UI Automation for FlyoutButton window.");
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"UI Automation error while checking button availability: {ex.Message}.");
                }
            }
            else
            {
                Debug.WriteLine("FlyoutButton window not found!");
            }

            // Check if coordinates are (0,0)
            if (tempTargetX == 0 && tempTargetY == 0)
            {
                Debug.WriteLine("Button is not available to be clicked (coordinates are (0,0)). Disabling monitoring and playing beep.");
                isMonitoring = false;
                toggleMonitoringMenuItem.Text = "Enable auto-GhostBust (Ctrl+Shift+Z)";
                Console.Beep(500, 500); // 500 Hz, 500 ms beep
                return false;
            }

            // Update global coordinates if button is available
            targetX = tempTargetX;
            targetY = tempTargetY;
            return true;
        }

        private static void ToggleMonitoring(object sender, EventArgs e)
        {
            isMonitoring = !isMonitoring;
            if (isMonitoring)
            {
                lastWindowId = GetForegroundWindow();
                string windowClass = GetWindowClassName(lastWindowId);
                if (windowClass == "Windows.UI.Core.CoreWindow" || windowClass == "Shell_TrayWnd" ||
                    windowClass == "SysListView32" || windowClass == "ClockFlyoutWindow" ||
                    windowClass == "TrayClockWClass" || windowClass == "TrayShowDesktopButtonWClass" ||
                    windowClass == "NotifyIconOverflowWindow")
                    lastWindowId = IntPtr.Zero;
                hasEvent = false;
                lastEventTime = 0;
                toggleMonitoringMenuItem.Text = "Disable auto-GhostBust (Ctrl+Shift+Z)";
                Debug.WriteLine("Monitoring toggled ON. Reset event state.");
            }
            else
            {
                hasEvent = false;
                lastEventTime = 0;
                toggleMonitoringMenuItem.Text = "Enable auto-GhostBust (Ctrl+Shift+Z)";
                Debug.WriteLine("Monitoring toggled OFF. Reset event state.");
            }
            Console.Beep(2000, 250); // Beep on toggle (2000 Hz, 250ms)
        }

        private static void RestartEInkPlus(object sender, EventArgs e)
        {
            Debug.WriteLine("Restarting EInkPlus processes...");

            try
            {
                Console.Beep(350, 200);

                // Step 1: Kill the processes "LenovoGen4.Launcher" and "LenovoGen4.FlyoutButton"
                Process[] launcherProcesses = Process.GetProcessesByName("LenovoGen4.Launcher");
                Process[] flyoutProcesses = Process.GetProcessesByName("LenovoGen4.FlyoutButton");

                foreach (Process process in launcherProcesses)
                {
                    Debug.WriteLine($"Killing process LenovoGen4.Launcher (PID: {process.Id})...");
                    process.Kill();
                    process.WaitForExit(5000); // Wait up to 5 seconds for the process to exit
                    Debug.WriteLine($"Process LenovoGen4.Launcher (PID: {process.Id}) terminated.");
                }

                foreach (Process process in flyoutProcesses)
                {
                    Debug.WriteLine($"Killing process LenovoGen4.FlyoutButton (PID: {process.Id})...");
                    process.Kill();
                    process.WaitForExit(5000); // Wait up to 5 seconds for the process to exit
                    Debug.WriteLine($"Process LenovoGen4.FlyoutButton (PID: {process.Id}) terminated.");
                }

                // Step 3: Wait until both processes are no longer running
                while (Process.GetProcessesByName("LenovoGen4.Launcher").Length > 0 ||
                       Process.GetProcessesByName("LenovoGen4.FlyoutButton").Length > 0)
                {
                    Debug.WriteLine("Waiting for LenovoGen4 processes to terminate...");
                    System.Threading.Thread.Sleep(100);
                }
                Debug.WriteLine("All LenovoGen4 processes have terminated.");

                // Step 4: Add a brief delay of 1000 ms
                System.Threading.Thread.Sleep(1000);
                Debug.WriteLine("Waited 1000 ms after process termination.");

                // Step 5: Restart LenovoGen4.Launcher.exe with proper working directory and window style
                string launcherPath = @"C:\Program Files\Lenovo\ThinkBookEinkPlus\LenovoGen4.Launcher.exe";
                Debug.WriteLine($"Starting LenovoGen4.Launcher from {launcherPath}...");

                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = launcherPath,
                    WorkingDirectory = @"C:\Program Files\Lenovo\ThinkBookEinkPlus", // Set working directory
                    WindowStyle = ProcessWindowStyle.Normal // Set to "Normal window"
                };

                Process.Start(startInfo);

                // Step 6: Wait until LenovoGen4.Launcher is running again
                while (Process.GetProcessesByName("LenovoGen4.Launcher").Length == 0)
                {
                    Debug.WriteLine("Waiting for LenovoGen4.Launcher to restart...");
                    System.Threading.Thread.Sleep(100);
                }
                Debug.WriteLine("LenovoGen4.Launcher has restarted.");

                // Step 7: Play beep sequence: 1000 Hz for 100 ms, sleep 100 ms, 2000 Hz for 100 ms
                Console.Beep(1000, 100);
                System.Threading.Thread.Sleep(100);
                Console.Beep(2000, 100);

                Debug.WriteLine("EInkPlus restart completed successfully.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error during EInkPlus restart: {ex.Message}");
                // Play error beep if something goes wrong: 200 Hz for 1000 ms
                Console.Beep(200, 1000);
            }
        }

        private static void PerformClickAndRestore(object sender, EventArgs e)
        {
            // Check if button is available before proceeding
            if (!IsButtonAvailable())
            {
                Debug.WriteLine("Button not available in PerformClickAndRestore. Action aborted.");
                return;
            }

            // Step 1: Beep (1000 Hz, 100 ms) - Commented out as per request
            // Console.Beep(1000, 100);
            Debug.WriteLine("Beep played (1000 Hz, 100 ms) - Commented out.");

            // Step 2: Record the current cursor position
            GetCursorPos(out POINT originalPos);
            originalX = originalPos.X;
            originalY = originalPos.Y;
            Debug.WriteLine($"Recorded original cursor position: ({originalX}, {originalY})");

            // Step 3: Record the foreground window
            originalWindowId = GetForegroundWindow();
            string originalWindowClass = GetWindowClassName(originalWindowId);
            Debug.WriteLine($"Recorded original foreground window: {originalWindowId} (class: {originalWindowClass})");

            // Step 4: Move the cursor to the ClickablePoint
            SetCursorPos(targetX, targetY);
            Debug.WriteLine($"Moved cursor to ClickablePoint: ({targetX}, {targetY})");

            // Step 5: Simulate a click
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            Debug.WriteLine("Simulated mouse click (left down and up).");

            // Step 6: Move the cursor back to its original position
            SetCursorPos(originalX, originalY);
            Debug.WriteLine($"Restored cursor position to: ({originalX}, {originalY})");

            // Step 7: Restore the foreground window, suppressing focus change events
            if (originalWindowId != IntPtr.Zero)
            {
                isRestoringFocus = true; // Set flag to suppress focus change events
                Debug.WriteLine("Setting isRestoringFocus to true to suppress focus change events.");
                SetForegroundWindow(originalWindowId);
                Debug.WriteLine($"Restored foreground window: {originalWindowId} (class: {originalWindowClass})");
                System.Threading.Thread.Sleep(100); // Short delay to debounce any transient focus changes
                isRestoringFocus = false; // Reset flag
                Debug.WriteLine("Reset isRestoringFocus to false after focus restoration.");
            }
            else
            {
                Debug.WriteLine("No original window to restore.");
            }
        }

        private static void ShowAboutDialog(object sender, EventArgs e)
        {
            MessageBox.Show(
                "GhostBustPlus, for all of your Lenovo e-Ink needs.\n" +
                "Designed for the ThinkBook Plus Gen 4 laptop.\n" +
                "Copyright (c) 2025 by joncox. All rights reserved.\n" +
                "No warranty is implied or fitness for any purpose.\n",
                "About GhostBustPlus",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private static void ExitApplication(object sender, EventArgs e)
        {
            Debug.WriteLine("Exiting application.");

            // Clean up UI Automation scroll event handler
            if (scrollEventHandler != null && currentScrollElement != null)
            {
                try
                {
                    Automation.RemoveAutomationPropertyChangedEventHandler(currentScrollElement, scrollEventHandler);
                    Debug.WriteLine("Removed scroll event handler on application exit.");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error removing scroll event handler on exit: {ex.Message}");
                }
            }

            UnhookWindowsHookEx(mouseHook);
            UnhookWindowsHookEx(keyboardHook);
            UnregisterGlobalHotKey();
            trayIcon.Visible = false;
            Application.Exit();
        }
    }
}