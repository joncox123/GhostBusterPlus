/*
 * GhostBusterPlus - Screen Refresh Utility for Windows 11
 *
 * Overview:
 * GhostBusterPlus is a system tray application designed to enhance the Lenovo ThinkBook Plus Gen 4 experience
 * by automatically refreshing the screen when significant changes are detected and user input has been idle.
 * The app captures periodic screenshots, compares them for changes, and simulates a function key press to refresh
 * the screen after a user-configurable delay. Settings are stored in an INI file.
 *
 * Structure:
 * The application is built using C# .NET Framework 4.8 with Windows Forms for the system tray interface. It uses timers to
 * monitor mouse movement, capture screenshots, and check refresh conditions. Low-level keyboard and mouse hooks
 * track user input. Screenshots are captured and processed using DirectX for GPU-accelerated performance with custom shaders.
 * The app includes a system tray icon with a context menu for configuration and an about dialog.
 *
 * Classes and Methods:
 *
 * ScreenRefreshApp (class):
 *   Main application class managing the system tray, timers, screenshots, and input monitoring.
 *   - ScreenRefreshApp(): Constructor, initializes components.
 *   - InitializeTrayIcon(): Sets up the system tray icon and context menu.
 *   - UpdateTrayMenu(): Updates menu checkmarks and text to reflect current settings.
 *   - LoadSettings(): Loads settings from the INI file.
 *   - SaveSettings(): Saves settings to the INI file.
 *   - InitializeTimers(): Initializes timers for mouse, screenshot, and refresh checks.
 *   - Shutdown(): Stops timers, releases hooks, and disposes resources before exiting.
 *   - InitializeInputHooks(): Sets up keyboard and mouse hooks.
 *   - CheckMouseInput(): Checks for mouse movement and updates timestamps.
 *   - TakeInitialScreenshot(): Captures an initial screenshot if enabled.
 *   - CaptureAndCompareScreenshotAsync(): Captures and compares screenshots, updates refresh flag.
 *   - CheckForRefresh(): Checks if a screen refresh should be performed.
 *   - RefreshScreen(): Simulates a key press and emits a beep.
 *   - Dispose(): Cleans up resources.
 *
 * KeyboardHook (static class):
 *   Handles low-level keyboard hooking for global key events.
 *   - SetHook(): Sets up the keyboard hook.
 *   - ReleaseHook(): Releases the keyboard hook.
 *   - ClearHandlers(): Clears all event handlers for KeyDown.
 *   - SetHook(proc): Configures the hook for the current process.
 *   - HookCallback(): Processes keyboard events.
 *   - AddHotkey(): Registers a hotkey combination (e.g., Shift+Ctrl+D).
 *
 * MouseHook (static class):
 *   Handles low-level mouse hooking for global button events.
 *   - SetHook(): Sets up the mouse hook.
 *   - ReleaseHook(): Releases the mouse hook.
 *   - ClearHandlers(): Clears all event handlers for MouseDown.
 *   - SetHook(proc): Configures the hook for the current process.
 *   - HookCallback(): Processes mouse button events.
 *
 * Program (static class):
 *   Entry point for the application.
 *   - Main(): Initializes and runs the application.
 */

using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace ScreenRefreshApp
{
    /// <summary>
    /// Main application class for the screen refresh utility.
    /// Manages system tray icon, timers, screenshot capture, and input monitoring.
    /// </summary>
    public class ScreenRefreshApp : System.Windows.Forms.Form
    {
        private System.Windows.Forms.NotifyIcon trayIcon; // System tray icon
        private System.Windows.Forms.ContextMenuStrip trayMenu; // Context menu for tray icon
        private System.Windows.Forms.Timer mouseCheckTimer; // Timer for checking mouse movement
        private System.Windows.Forms.Timer screenshotTimer; // Timer for capturing screenshots
        private System.Windows.Forms.Timer refreshCheckTimer; // Timer for checking refresh conditions
        private System.Drawing.Point lastMousePosition; // Last recorded mouse cursor position
        private long lastMouseInputTime; // Timestamp (ms) of last mouse movement
        private long lastButtonInputTime; // Timestamp (ms) of last keyboard/mouse button input
        private long lastScreenChangeTime; // Timestamp (ms) of last significant screen change
        private bool doRefresh; // Flag indicating if a screen refresh is needed
        private bool screenshotsEnabled; // Whether screenshot capture is enabled
        private int screenshotPeriodMs; // Screenshot capture interval (ms)
        private int userInputDelayMs; // Delay after user input before refresh (ms)
        private System.Windows.Forms.Keys refreshKey; // Key to simulate for screen refresh
        private string iniPath = System.IO.Path.Combine(System.Windows.Forms.Application.StartupPath, "settings.ini"); // Path to settings INI file
        private const double MOUSE_CHECK_TIMER_SEC = 0.2; // Mouse check interval (seconds)
        private readonly System.Threading.CancellationTokenSource cancellationTokenSource = new System.Threading.CancellationTokenSource();
        private readonly Processor screenshotProcessor; // DirectX processor
        private double refreshThresholdPct = 3.0; // Default to 3%
        private bool firstRunMessageShown = false; // Flag to track if first run message has been shown
        
        // P/Invoke for simulating keyboard events
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, int dwExtraInfo);
        private const int KEYEVENTF_KEYUP = 0x0002;

        // P/Invoke for generating console beep
        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        private static extern void Beep(uint dwFreq, uint dwDuration);

        /// <summary>
        /// Constructor for the ScreenRefreshApp.
        /// Initializes tray icon, loads settings, sets up timers, and registers input hooks.
        /// </summary>
        public ScreenRefreshApp()
        {
            this.Visible = false;
            this.ShowInTaskbar = false;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Size = new System.Drawing.Size(0, 0);

            screenshotProcessor = new Processor();
            InitializeTrayIcon();
            LoadSettings();
            lastScreenChangeTime = System.Environment.TickCount; // Initialize to now
            
            // Show the first run message if it hasn't been shown before
            if (!firstRunMessageShown)
            {
                ShowFirstRunMessage();
            }
            
            UpdateTrayMenu();
            InitializeTimers();
            InitializeInputHooks();

            // Add a message filter to catch mouse wheel events at the application level
            Application.AddMessageFilter(new GlobalMouseWheelMessageFilter(() => lastButtonInputTime = Environment.TickCount));

            TakeInitialScreenshot();

            KeyboardHook.AddHotkey(System.Windows.Forms.Keys.D | System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift, () =>
            {
                screenshotsEnabled = !screenshotsEnabled;
                UpdateTrayMenu();
                SaveSettings();
                screenshotTimer.Enabled = screenshotsEnabled;
            });

            // Add the global mouse wheel message filter
            Application.AddMessageFilter(new GlobalMouseWheelMessageFilter(() =>
            {
                lastButtonInputTime = Environment.TickCount;
            }));
        }

        /// <summary>
        /// Displays a welcome message with setup instructions and privacy information on first run.
        /// </summary>
        private void ShowFirstRunMessage()
        {
            System.Windows.Forms.MessageBox.Show(
                "GhostBusterPlus works by monitoring pixel changes on your screen, and automatically pressing a function key " +
                "(e.g. F4, etc.) if the screen changes exceed a percentage threshold. To work, you must be using the latest " +
                "version of Lenovo EInkPlus and enable the clear ghosts shortcut key in the EInkPlus settings to set the key to F4.\n\n" +
                "A note regarding privacy and security: This app does not have network connectivity, elevated privileges or recording " +
                "features. In fact, the screen shot is never transferred into the program, but rather, all display processing is done " +
                "in DirectX on the GPU, and only the percentage pixel change is returned to the program.",
                "Welcome to GhostBusterPlus",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
            
            // Mark the message as shown and save this setting
            firstRunMessageShown = true;
            SaveSettings();
        }

        /// <summary>
        /// Initializes the system tray icon and context menu.
        /// </summary>
        private void InitializeTrayIcon()
        {
            trayMenu = new System.Windows.Forms.ContextMenuStrip();

            // Screenshot period sub-menu
            System.Windows.Forms.ToolStripMenuItem screenshotPeriodMenu = new System.Windows.Forms.ToolStripMenuItem("Screenshot Period");
            int[] periods = { 250, 500, 1000, 2000, 4000 };
            foreach (int period in periods)
            {
                System.Windows.Forms.ToolStripMenuItem item = new System.Windows.Forms.ToolStripMenuItem($"{period} ms");
                item.Checked = period == screenshotPeriodMs;
                item.Click += (s, e) =>
                {
                    screenshotPeriodMs = period;
                    screenshotTimer.Interval = period;
                    UpdateTrayMenu();
                    SaveSettings();
                };
                screenshotPeriodMenu.DropDownItems.Add(item);
            }

            // Enable/disable screenshots menu item
            System.Windows.Forms.ToolStripMenuItem enableScreenshotsMenu = new System.Windows.Forms.ToolStripMenuItem(screenshotsEnabled ? "Enabled" : "Disabled");
            enableScreenshotsMenu.Click += (s, e) =>
            {
                screenshotsEnabled = !screenshotsEnabled;
                screenshotTimer.Enabled = screenshotsEnabled;
                UpdateTrayMenu();
                SaveSettings();
            };

            // Refresh key sub-menu
            System.Windows.Forms.ToolStripMenuItem refreshKeyMenu = new System.Windows.Forms.ToolStripMenuItem("Refresh Key");
            System.Windows.Forms.Keys[] keys = { System.Windows.Forms.Keys.F4, System.Windows.Forms.Keys.F5, System.Windows.Forms.Keys.F8, System.Windows.Forms.Keys.F9 };
            foreach (System.Windows.Forms.Keys key in keys)
            {
                System.Windows.Forms.ToolStripMenuItem item = new System.Windows.Forms.ToolStripMenuItem(key.ToString());
                item.Checked = key == refreshKey;
                item.Click += (s, e) =>
                {
                    refreshKey = key;
                    UpdateTrayMenu();
                    SaveSettings();
                };
                refreshKeyMenu.DropDownItems.Add(item);
            }

            // User input delay sub-menu
            System.Windows.Forms.ToolStripMenuItem inputDelayMenu = new System.Windows.Forms.ToolStripMenuItem("User Input Delay");
            int[] delays = { 250, 500, 1000, 2000, 4000, 8000 };
            foreach (int delay in delays)
            {
                System.Windows.Forms.ToolStripMenuItem item = new System.Windows.Forms.ToolStripMenuItem($"{delay} ms");
                item.Checked = delay == userInputDelayMs;
                item.Click += (s, e) =>
                {
                    userInputDelayMs = delay;
                    refreshCheckTimer.Interval = userInputDelayMs / 4;
                    UpdateTrayMenu();
                    SaveSettings();
                };
                inputDelayMenu.DropDownItems.Add(item);
            }

            // Refresh threshold sub-menu
            System.Windows.Forms.ToolStripMenuItem thresholdMenu = new System.Windows.Forms.ToolStripMenuItem("Refresh Threshold");
            double[] thresholds = { 1.0, 2.0, 3.0, 5.0, 10.0, 15.0, 20.0 };
            foreach (double threshold in thresholds)
            {
                System.Windows.Forms.ToolStripMenuItem item = new System.Windows.Forms.ToolStripMenuItem($"{threshold}%");
                item.Checked = threshold == refreshThresholdPct;
                item.Click += (s, e) =>
                {
                    refreshThresholdPct = threshold;
                    screenshotProcessor.PixelThresholdPct = threshold;
                    UpdateTrayMenu();
                    SaveSettings();
                };
                thresholdMenu.DropDownItems.Add(item);
            }

            // About menu item
            System.Windows.Forms.ToolStripMenuItem aboutMenu = new System.Windows.Forms.ToolStripMenuItem("About GhostBusterPlus...");
            aboutMenu.Click += (s, e) =>
            {
                System.Windows.Forms.MessageBox.Show(
                    "GhostBusterPlus v0.2, by joncox123. Enhancing your Lenovo ThinkBook Plus Gen 4 experience. " +
                    "Copyright (c) 2025, all rights reserved. No warranty or suitability for any purpose is implied or provided.",
                    "About GhostBusterPlus",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
            };

            // Exit menu item
            System.Windows.Forms.ToolStripMenuItem exitMenu = new System.Windows.Forms.ToolStripMenuItem("Exit");
            exitMenu.Click += (s, e) =>
            {
                Shutdown();
                System.Windows.Forms.Application.Exit();
                System.Environment.Exit(0);
            };

            trayMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] 
            { 
                screenshotPeriodMenu, 
                enableScreenshotsMenu, 
                refreshKeyMenu, 
                inputDelayMenu, 
                thresholdMenu, 
                aboutMenu, 
                exitMenu 
            });

            trayIcon = new System.Windows.Forms.NotifyIcon
            {
                Icon = new System.Drawing.Icon(System.IO.Path.Combine(System.Windows.Forms.Application.StartupPath, "icon.ico")),
                ContextMenuStrip = trayMenu,
                Visible = true,
                Text = "GhostBusterPlus"
            };
        }

        /// <summary>
        /// Updates the tray menu to reflect current settings.
        /// </summary>
        private void UpdateTrayMenu()
        {
            foreach (System.Windows.Forms.ToolStripMenuItem item in trayMenu.Items)
            {
                if (item.Text == "Screenshot Period")
                {
                    foreach (System.Windows.Forms.ToolStripMenuItem subItem in item.DropDownItems)
                    {
                        bool isSelected = subItem.Text == $"{screenshotPeriodMs} ms";
                        subItem.Checked = isSelected;
                        subItem.Font = new System.Drawing.Font(subItem.Font ?? System.Drawing.SystemFonts.MenuFont, isSelected ? System.Drawing.FontStyle.Bold : System.Drawing.FontStyle.Regular);
                    }
                }
                else if (item.Text == "Enabled" || item.Text == "Disabled")
                {
                    item.Text = screenshotsEnabled ? "Enabled" : "Disabled";
                }
                else if (item.Text == "Refresh Key")
                {
                    foreach (System.Windows.Forms.ToolStripMenuItem subItem in item.DropDownItems)
                    {
                        bool isSelected = subItem.Text == refreshKey.ToString();
                        subItem.Checked = isSelected;
                        subItem.Font = new System.Drawing.Font(subItem.Font ?? System.Drawing.SystemFonts.MenuFont, isSelected ? System.Drawing.FontStyle.Bold : System.Drawing.FontStyle.Regular);
                    }
                }
                else if (item.Text == "User Input Delay")
                {
                    foreach (System.Windows.Forms.ToolStripMenuItem subItem in item.DropDownItems)
                    {
                        bool isSelected = subItem.Text == $"{userInputDelayMs} ms";
                        subItem.Checked = isSelected;
                        subItem.Font = new System.Drawing.Font(subItem.Font ?? System.Drawing.SystemFonts.MenuFont, isSelected ? System.Drawing.FontStyle.Bold : System.Drawing.FontStyle.Regular);
                    }
                }
                else if (item.Text == "Refresh Threshold")
                {
                    foreach (System.Windows.Forms.ToolStripMenuItem subItem in item.DropDownItems)
                    {
                        bool isSelected = subItem.Text == $"{refreshThresholdPct}%";
                        subItem.Checked = isSelected;
                        subItem.Font = new System.Drawing.Font(subItem.Font ?? System.Drawing.SystemFonts.MenuFont, isSelected ? System.Drawing.FontStyle.Bold : System.Drawing.FontStyle.Regular);
                    }
                }
            }
        }

        /// <summary>
        /// Loads settings from the INI file.
        /// </summary>
        private void LoadSettings()
        {
            screenshotsEnabled = true;
            screenshotPeriodMs = 500;
            userInputDelayMs = 4000; // Default to 4000 ms
            refreshKey = System.Windows.Forms.Keys.F4;
            refreshThresholdPct = 3.0;
            firstRunMessageShown = false; // Default is false (show message)
            
            if (System.IO.File.Exists(iniPath))
            {
                var lines = System.IO.File.ReadAllLines(iniPath);
                foreach (var line in lines)
                {
                    var parts = line.Split('=');
                    if (parts.Length == 2)
                    {
                        switch (parts[0].Trim())
                        {
                            case "ScreenshotsEnabled": screenshotsEnabled = bool.Parse(parts[1].Trim()); break;
                            case "ScreenshotPeriodMs": screenshotPeriodMs = int.Parse(parts[1].Trim()); break;
                            case "UserInputDelayMs": userInputDelayMs = int.Parse(parts[1].Trim()); break;
                            case "RefreshKey": refreshKey = (System.Windows.Forms.Keys)System.Enum.Parse(typeof(System.Windows.Forms.Keys), parts[1].Trim()); break;
                            case "RefreshThresholdPct": refreshThresholdPct = double.Parse(parts[1].Trim()); break;
                            case "FirstRunMessageShown": firstRunMessageShown = bool.Parse(parts[1].Trim()); break;
                        }
                    }
                }
            }
            
            // Apply the threshold setting to the processor
            screenshotProcessor.PixelThresholdPct = refreshThresholdPct;
        }

        /// <summary>
        /// Saves settings to the INI file.
        /// </summary>
        private void SaveSettings()
        {
            var settings = new[]
            {
                $"ScreenshotsEnabled={screenshotsEnabled}",
                $"ScreenshotPeriodMs={screenshotPeriodMs}",
                $"UserInputDelayMs={userInputDelayMs}",
                $"RefreshKey={refreshKey}",
                $"RefreshThresholdPct={refreshThresholdPct}",
                $"FirstRunMessageShown={firstRunMessageShown}"
            };
            System.IO.File.WriteAllLines(iniPath, settings);
        }

        /// <summary>
        /// Initializes all timers.
        /// </summary>
        private void InitializeTimers()
        {
            mouseCheckTimer = new System.Windows.Forms.Timer { Interval = (int)(MOUSE_CHECK_TIMER_SEC * 1000) };
            mouseCheckTimer.Tick += (s, e) => CheckMouseInput();
            mouseCheckTimer.Start();

            screenshotTimer = new System.Windows.Forms.Timer { Interval = screenshotPeriodMs };
            screenshotTimer.Tick += async (s, e) => await CaptureAndCompareScreenshotAsync();
            screenshotTimer.Enabled = screenshotsEnabled;

            refreshCheckTimer = new System.Windows.Forms.Timer { Interval = 100 };
            refreshCheckTimer.Tick += (s, e) => CheckForRefresh();
            refreshCheckTimer.Start();
        }

        /// <summary>
        /// Shuts down the application cleanly.
        /// </summary>
        private void Shutdown()
        {
            // Signal cancellation to stop async tasks
            cancellationTokenSource.Cancel();

            // Stop timers
            mouseCheckTimer?.Stop();
            screenshotTimer?.Stop();
            refreshCheckTimer?.Stop();

            // Release hooks
            KeyboardHook.ReleaseHook();
            MouseHook.ReleaseHook();
            KeyboardHook.ClearHandlers();
            MouseHook.ClearHandlers();

            // Dispose tray icon
            if (trayIcon != null)
            {
                trayIcon.Visible = false;
                trayIcon.Dispose();
                trayIcon = null;
                System.Threading.Thread.Sleep(100); // Ensure system tray updates
            }

            // Dispose other resources
            trayMenu?.Dispose();
            mouseCheckTimer?.Dispose();
            screenshotTimer?.Dispose();
            refreshCheckTimer?.Dispose();
            screenshotProcessor?.Dispose();

            System.Windows.Forms.Application.Exit();
            System.Environment.Exit(0);
        }

        /// <summary>
        /// Initializes keyboard and mouse hooks.
        /// </summary>
        private void InitializeInputHooks()
        {
            KeyboardHook.SetHook();
            MouseHook.SetHook();
            KeyboardHook.KeyDown += (s, e) => lastButtonInputTime = System.Environment.TickCount;
            MouseHook.MouseDown += (s, e) => lastButtonInputTime = System.Environment.TickCount;
        }

        /// <summary>
        /// Captures an initial screenshot if enabled.
        /// </summary>
        private void TakeInitialScreenshot()
        {
            if (screenshotsEnabled)
            {
                screenshotProcessor.ProcessScreenshotOnGPU();
            }
        }

        /// <summary>
        /// Captures a new screenshot, compares it with the previous one, and updates doRefresh.
        /// </summary>
        private async System.Threading.Tasks.Task CaptureAndCompareScreenshotAsync()
        {
            if (!screenshotsEnabled) return;

            try
            {
                bool significantChange = await System.Threading.Tasks.Task.Run(() => screenshotProcessor.ProcessScreenshotOnGPU(), cancellationTokenSource.Token);
                if (significantChange)
                {
                    doRefresh = true;
                    lastScreenChangeTime = System.Environment.TickCount; // Record when the change was detected
                }
            }
            catch (System.Threading.Tasks.TaskCanceledException)
            {
                System.Diagnostics.Debug.WriteLine("Screenshot capture task was cancelled during shutdown.");
            }
        }

        /// <summary>
        /// Checks mouse movement and updates timestamps.
        /// </summary>
        private void CheckMouseInput()
        {
            var currentPosition = System.Windows.Forms.Cursor.Position;
            if (currentPosition != lastMousePosition)
            {
                lastMouseInputTime = System.Environment.TickCount;
                lastMousePosition = currentPosition;
            }
        }

        /// <summary>
        /// Checks if a screen refresh should be performed.
        /// </summary>
        private void CheckForRefresh()
        {
            long currentTime = System.Environment.TickCount;
            if (doRefresh &&
                (currentTime - lastMouseInputTime >= userInputDelayMs) &&
                (currentTime - lastButtonInputTime >= userInputDelayMs) &&
                (currentTime - lastScreenChangeTime >= userInputDelayMs)) // Wait for screen to stabilize
            {
                RefreshScreen();
                doRefresh = false;
            }
        }

        /// <summary>
        /// Simulates a key press to refresh the screen and emits a beep.
        /// </summary>
        private void RefreshScreen()
        {
            keybd_event((byte)refreshKey, 0, 0, 0);
            System.Threading.Thread.Sleep(10);
            keybd_event((byte)refreshKey, 0, KEYEVENTF_KEYUP, 0);
            Beep(1000, 200);
        }

        /// <summary>
        /// Cleans up resources.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                cancellationTokenSource.Dispose();
                trayIcon?.Dispose();
                trayMenu?.Dispose();
                mouseCheckTimer?.Dispose();
                screenshotTimer?.Dispose();
                refreshCheckTimer?.Dispose();
                screenshotProcessor?.Dispose();
                KeyboardHook.ReleaseHook();
                MouseHook.ReleaseHook();
            }
            base.Dispose(disposing);
        }
    }

    /// <summary>
    /// Static class for handling low-level keyboard hooks.
    /// </summary>
    public static class KeyboardHook
    {
        private static System.IntPtr hookId = System.IntPtr.Zero;
        private static LowLevelKeyboardProc proc = HookCallback;
        public static event System.EventHandler<System.Windows.Forms.KeyEventArgs> KeyDown;

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern System.IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, System.IntPtr hMod, uint dwThreadId);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool UnhookWindowsHookEx(System.IntPtr hhk);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern System.IntPtr CallNextHookEx(System.IntPtr hhk, int nCode, System.IntPtr wParam, System.IntPtr lParam);

        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        private static extern System.IntPtr GetModuleHandle(string lpModuleName);

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;

        private delegate System.IntPtr LowLevelKeyboardProc(int nCode, System.IntPtr wParam, System.IntPtr lParam);

        public static void SetHook()
        {
            if (hookId == System.IntPtr.Zero)
            {
                using (var curProcess = System.Diagnostics.Process.GetCurrentProcess())
                using (var curModule = curProcess.MainModule)
                {
                    hookId = SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
                }
            }
        }

        public static void ReleaseHook()
        {
            if (hookId != System.IntPtr.Zero)
            {
                UnhookWindowsHookEx(hookId);
                hookId = System.IntPtr.Zero;
            }
        }

        public static void ClearHandlers()
        {
            if (KeyDown != null)
            {
                foreach (System.Delegate d in KeyDown.GetInvocationList())
                {
                    KeyDown -= (System.EventHandler<System.Windows.Forms.KeyEventArgs>)d;
                }
            }
        }

        private static System.IntPtr HookCallback(int nCode, System.IntPtr wParam, System.IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (System.IntPtr)WM_KEYDOWN)
            {
                int vkCode = System.Runtime.InteropServices.Marshal.ReadInt32(lParam);
                KeyDown?.Invoke(null, new System.Windows.Forms.KeyEventArgs((System.Windows.Forms.Keys)vkCode));
            }
            return CallNextHookEx(hookId, nCode, wParam, lParam);
        }

        public static void AddHotkey(System.Windows.Forms.Keys key, System.Action action)
        {
            bool ctrl = (key & System.Windows.Forms.Keys.Control) != 0;
            bool shift = (key & System.Windows.Forms.Keys.Shift) != 0;
            System.Windows.Forms.Keys keyCode = key & ~(System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift);
            KeyboardHook.KeyDown += (s, e) =>
            {
                if (System.Windows.Forms.Control.ModifierKeys.HasFlag(System.Windows.Forms.Keys.Control) == ctrl &&
                    System.Windows.Forms.Control.ModifierKeys.HasFlag(System.Windows.Forms.Keys.Shift) == shift &&
                    e.KeyCode == keyCode)
                {
                    action();
                }
            };
        }
    }

    /// <summary>
    /// Static class for handling low-level mouse hooks.
    /// </summary>
    public static class MouseHook
    {
        private static System.IntPtr hookId = System.IntPtr.Zero;
        private static LowLevelMouseProc proc = HookCallback;
        public static event System.EventHandler<System.Windows.Forms.MouseEventArgs> MouseDown;

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern System.IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, System.IntPtr hMod, uint dwThreadId);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool UnhookWindowsHookEx(System.IntPtr hhk);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern System.IntPtr CallNextHookEx(System.IntPtr hhk, int nCode, System.IntPtr wParam, System.IntPtr lParam);

        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        private static extern System.IntPtr GetModuleHandle(string lpModuleName);

        private const int WH_MOUSE_LL = 14;
        private const int WM_LBUTTONDOWN = 0x0201;
        private const int WM_RBUTTONDOWN = 0x0204;
        private const int WM_MBUTTONDOWN = 0x0207;
        private const int WM_XBUTTONDOWN = 0x020B;
        private const int WM_MOUSEWHEEL = 0x020A;
        private const int WM_MOUSEHWHEEL = 0x020E;

        private delegate System.IntPtr LowLevelMouseProc(int nCode, System.IntPtr wParam, System.IntPtr lParam);

        public static void SetHook()
        {
            if (hookId == System.IntPtr.Zero)
            {
                using (var curProcess = System.Diagnostics.Process.GetCurrentProcess())
                using (var curModule = curProcess.MainModule)
                {
                    hookId = SetWindowsHookEx(WH_MOUSE_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
                }
            }
        }

        public static void ReleaseHook()
        {
            if (hookId != System.IntPtr.Zero)
            {
                UnhookWindowsHookEx(hookId);
                hookId = System.IntPtr.Zero;
            }
        }

        public static void ClearHandlers()
        {
            if (MouseDown != null)
            {
                foreach (System.Delegate d in MouseDown.GetInvocationList())
                {
                    MouseDown -= (System.EventHandler<System.Windows.Forms.MouseEventArgs>)d;
                }
            }
        }

        private static System.IntPtr HookCallback(int nCode, System.IntPtr wParam, System.IntPtr lParam)
        {
            if (nCode >= 0)
            {
                int msg = wParam.ToInt32();
                MSLLHOOKSTRUCT hookStruct = (MSLLHOOKSTRUCT)System.Runtime.InteropServices.Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));
                System.Windows.Forms.MouseButtons button = System.Windows.Forms.MouseButtons.None;
                int clicks = 1;
                int delta = 0;

                switch (msg)
                {
                    case WM_LBUTTONDOWN:
                        button = System.Windows.Forms.MouseButtons.Left;
                        break;
                    case WM_RBUTTONDOWN:
                        button = System.Windows.Forms.MouseButtons.Right;
                        break;
                    case WM_MBUTTONDOWN:
                        button = System.Windows.Forms.MouseButtons.Middle;
                        break;
                    case WM_XBUTTONDOWN:
                        // XBUTTON1 = 0x0001, XBUTTON2 = 0x0002
                        button = ((hookStruct.mouseData >> 16) & 0xFFFF) == 1
                            ? System.Windows.Forms.MouseButtons.XButton1
                            : System.Windows.Forms.MouseButtons.XButton2;
                        break;
                    case WM_MOUSEWHEEL:
                        button = System.Windows.Forms.MouseButtons.None;
                        delta = (short)((hookStruct.mouseData >> 16) & 0xFFFF);
                        break;
                    case WM_MOUSEHWHEEL:
                        button = System.Windows.Forms.MouseButtons.None;
                        delta = (short)((hookStruct.mouseData >> 16) & 0xFFFF);
                        break;
                }

                // Fire MouseDown for all button presses and wheel events
                if (button != System.Windows.Forms.MouseButtons.None ||
                    msg == WM_MOUSEWHEEL || msg == WM_MOUSEHWHEEL)
                {
                    MouseDown?.Invoke(null, new System.Windows.Forms.MouseEventArgs(
                        button, clicks, hookStruct.pt.X, hookStruct.pt.Y, delta));
                }
            }
            return CallNextHookEx(hookId, nCode, wParam, lParam);
        }

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct MSLLHOOKSTRUCT
        {
            public POINT pt;
            public uint mouseData;
            public uint flags;
            public uint time;
            public System.IntPtr dwExtraInfo;
        }
    }

    // Message filter to catch mouse wheel and horizontal wheel events globally
    public class GlobalMouseWheelMessageFilter : IMessageFilter
    {
        private readonly Action onWheel;

        // Windows message constants
        private const int WM_MOUSEWHEEL = 0x020A;
        private const int WM_MOUSEHWHEEL = 0x020E;

        public GlobalMouseWheelMessageFilter(Action onWheel)
        {
            this.onWheel = onWheel;
        }

        public bool PreFilterMessage(ref Message m)
        {
            if (m.Msg == WM_MOUSEWHEEL || m.Msg == WM_MOUSEHWHEEL)
            {
                onWheel?.Invoke();
            }
            return false; // Do not block the message
        }
    }

    /// <summary>
    /// Program entry point.
    /// </summary>
    static class Program
    {
        [System.STAThread]
        static void Main()
        {
            System.Windows.Forms.Application.EnableVisualStyles();
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
            using (var app = new ScreenRefreshApp())
            {
                System.Windows.Forms.Application.Run();
            }
        }
    }
}