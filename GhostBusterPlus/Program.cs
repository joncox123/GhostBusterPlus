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
using System.Diagnostics;
using System.Text;

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
        private string userThemesPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Windows", "Themes");
        private string eInkThemePath;
        private string darkThemePath;
        private bool detectCursorMovement = false; // Default to false/disabled
        
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
            
            // Initialize theme paths - MUST happen regardless of first run status
            InitializeThemePaths();
            
            // Copy theme files on first run only
            if (!firstRunMessageShown)
            {
                CopyThemeFiles();
                ShowFirstRunMessage();
            }
            
            UpdateTrayMenu();
            InitializeTimers();
            InitializeInputHooks();

            // Add a message filter to catch mouse wheel events at the application level
            Application.AddMessageFilter(new GlobalMouseWheelMessageFilter(() => lastButtonInputTime = Environment.TickCount));

            TakeInitialScreenshot();

            // Add hotkeys for theme switching
            KeyboardHook.AddHotkey(System.Windows.Forms.Keys.D | System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift, () => 
            {
                RestartEInkPlus(null, null); // Use RestartEInkPlus method for Ctrl+Shift+D
            });
            
            KeyboardHook.AddHotkey(System.Windows.Forms.Keys.X | System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift, () =>
            {
                ToggleScreenshots(); // Keep Ctrl+Shift+X for toggling screenshots
            });
            
            KeyboardHook.AddHotkey(System.Windows.Forms.Keys.S | System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift, () =>
            {
                ApplyTheme(darkThemePath);
            });
            
            KeyboardHook.AddHotkey(System.Windows.Forms.Keys.Z | System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift, () =>
            {
                ApplyTheme(eInkThemePath);
            });
        }

        /// <summary>
        /// Initializes theme paths regardless of whether it's first run or not
        /// </summary>
        private void InitializeThemePaths()
        {
            // Always set up the theme paths
            eInkThemePath = Path.Combine(userThemesPath, "eInk.theme");
            darkThemePath = Path.Combine(userThemesPath, "Dark.theme");
            
            // Verify theme files exist
            if (!File.Exists(eInkThemePath))
                System.Console.WriteLine($"Warning: eInk theme file not found at {eInkThemePath}");
            else
                System.Console.WriteLine($"eInk theme file found at {eInkThemePath}");
                
            if (!File.Exists(darkThemePath))
                System.Console.WriteLine($"Warning: Dark theme file not found at {darkThemePath}");
            else
                System.Console.WriteLine($"Dark theme file found at {darkThemePath}");
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

            // Enable/disable screenshots menu item - updated with hotkey in text
            System.Windows.Forms.ToolStripMenuItem enableScreenshotsMenu = new System.Windows.Forms.ToolStripMenuItem(screenshotsEnabled ? "Enabled (Ctrl+Shift+X)" : "Disabled (Ctrl+Shift+X)");
            enableScreenshotsMenu.Click += (s, e) =>
            {
                ToggleScreenshots();
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

            // Theme switching sub-menu
            System.Windows.Forms.ToolStripMenuItem themeMenu = new System.Windows.Forms.ToolStripMenuItem("Switch Theme");

            // Dark theme menu item
            System.Windows.Forms.ToolStripMenuItem darkThemeMenuItem = new System.Windows.Forms.ToolStripMenuItem("Dark (Ctrl+Shift+S)");
            darkThemeMenuItem.Click += (s, e) => ApplyTheme(darkThemePath);
            themeMenu.DropDownItems.Add(darkThemeMenuItem);

            // eInk theme menu item
            System.Windows.Forms.ToolStripMenuItem eInkThemeMenuItem = new System.Windows.Forms.ToolStripMenuItem("eInk (Ctrl+Shift+Z)");
            eInkThemeMenuItem.Click += (s, e) => ApplyTheme(eInkThemePath);
            themeMenu.DropDownItems.Add(eInkThemeMenuItem);

            // Restart EInkPlus menu item
            System.Windows.Forms.ToolStripMenuItem restartEInkPlusMenu = new System.Windows.Forms.ToolStripMenuItem("Restart EInkPlus (Ctrl+Shift+D)");
            restartEInkPlusMenu.Click += RestartEInkPlus;

            // Detect cursor movement menu item
            System.Windows.Forms.ToolStripMenuItem detectCursorMenu = new System.Windows.Forms.ToolStripMenuItem("Detect cursor movement");
            detectCursorMenu.Checked = detectCursorMovement;
            detectCursorMenu.Click += (s, e) =>
            {
                detectCursorMovement = !detectCursorMovement;
                UpdateTrayMenu();
                SaveSettings();
            };

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
                detectCursorMenu,      // New item
                restartEInkPlusMenu,   // New item
                refreshKeyMenu, 
                inputDelayMenu, 
                thresholdMenu, 
                themeMenu,
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

        /// <summary>
        /// Toggles screenshot capture on/off
        /// </summary>
        private void ToggleScreenshots()
        {
            screenshotsEnabled = !screenshotsEnabled;
            screenshotTimer.Enabled = screenshotsEnabled;
            UpdateTrayMenu();
            SaveSettings();
            
            // Display a notification
            trayIcon.ShowBalloonTip(
                1500,
                "Screenshot Capture",
                $"Screenshot capture {(screenshotsEnabled ? "enabled" : "disabled")}",
                ToolTipIcon.Info);
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
                else if (item.Text.Contains("Enabled") || item.Text.Contains("Disabled"))
                {
                    item.Text = screenshotsEnabled ? "Enabled (Ctrl+Shift+X)" : "Disabled (Ctrl+Shift+X)";
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
                else if (item.Text == "Detect cursor movement")
                {
                    item.Checked = detectCursorMovement;
                    item.Font = new System.Drawing.Font(item.Font ?? System.Drawing.SystemFonts.MenuFont, 
                        detectCursorMovement ? System.Drawing.FontStyle.Bold : System.Drawing.FontStyle.Regular);
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
            userInputDelayMs = 2000; // Default to 2000 ms
            refreshKey = System.Windows.Forms.Keys.F4;
            refreshThresholdPct = 3.0;
            firstRunMessageShown = false; // Default is false (show message)
            detectCursorMovement = false; // Default to disabled
            
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
                            case "DetectCursorMovement": detectCursorMovement = bool.Parse(parts[1].Trim()); break;
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
                $"FirstRunMessageShown={firstRunMessageShown}",
                $"DetectCursorMovement={detectCursorMovement}"
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
            MouseHook.MouseDown += (s, e) => 
            {
                // Explicitly check for wheel events as well as button presses
                lastButtonInputTime = System.Environment.TickCount;
                System.Console.WriteLine($"Mouse input: Button={e.Button}, Delta={e.Delta}, X={e.X}, Y={e.Y}");
            };

            // Add application-level message handling for touchpad gestures
            Application.AddMessageFilter(new GlobalMouseWheelMessageFilter(() =>
            {
                lastButtonInputTime = System.Environment.TickCount;
                System.Console.WriteLine("Mouse wheel message detected through message filter");
            }));
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
            if (!detectCursorMovement) return; // Skip if cursor detection is disabled
            
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
            // Beep(1000, 200);
        }

        /// <summary>
        /// Copies and prepares theme files for Windows 11 compatibility
        /// </summary>
        private void CopyThemeFiles()
        {
            try
            {
                // Create the themes directory if it doesn't exist
                Directory.CreateDirectory(userThemesPath);

                string appPath = Application.StartupPath;
                string sourceEInkPath = Path.Combine(appPath, "eInk.theme");
                string sourceDarkPath = Path.Combine(appPath, "Dark.theme");

                // Use the EXACT names as the source files - this is critical
                eInkThemePath = Path.Combine(userThemesPath, "eInk.theme");
                darkThemePath = Path.Combine(userThemesPath, "Dark.theme");

                if (File.Exists(sourceEInkPath))
                {
                    File.Copy(sourceEInkPath, eInkThemePath, true);

                    // Set appropriate file attributes
                    File.SetAttributes(eInkThemePath, FileAttributes.Normal);
                    System.Console.WriteLine($"Copied {sourceEInkPath} to {eInkThemePath}");
                }
                else
                {
                    System.Console.WriteLine($"Source file not found: {sourceEInkPath}");
                }

                if (File.Exists(sourceDarkPath))
                {
                    File.Copy(sourceDarkPath, darkThemePath, true);

                    // Set appropriate file attributes
                    File.SetAttributes(darkThemePath, FileAttributes.Normal);
                    System.Console.WriteLine($"Copied {sourceDarkPath} to {darkThemePath}");
                }
                else
                {
                    System.Console.WriteLine($"Source file not found: {sourceDarkPath}");
                }

                System.Console.WriteLine($"Theme files copied to {userThemesPath}");
            }
            catch (Exception ex)
            {
                System.Console.WriteLine($"Failed to copy theme files: {ex.Message}");
            }
        }

        /// <summary>
        /// Applies Windows theme and closes any opened settings windows
        /// </summary>
        /// <param name="themePath">Full path to the theme file</param>
        private void ApplyTheme(string themePath)
        {
            if (string.IsNullOrEmpty(themePath) || !File.Exists(themePath))
            {
                System.Console.WriteLine($"Theme file not found: {themePath}");
                return;
            }

            try
            {
                System.Console.WriteLine($"Applying theme: {themePath}");
                Process settingsProcess = null;
                Process themeProcess = null;

                // Method 1: Direct theme file execution (most effective)
                themeProcess = Process.Start(new ProcessStartInfo
                {
                    FileName = themePath,
                    UseShellExecute = true,
                    Verb = "open"
                });
                
                // Short wait to ensure the theme file is processed
                Thread.Sleep(800);
                
                // Method 2: Set registry value (for persistence)
                Microsoft.Win32.Registry.SetValue(
                    @"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes",
                    "CurrentTheme", 
                    themePath,
                    Microsoft.Win32.RegistryValueKind.String);

                // Create a timer to find and close Settings windows after a delay
                System.Windows.Forms.Timer cleanupTimer = new System.Windows.Forms.Timer();
                cleanupTimer.Interval = 1500; // Wait 1.5 seconds
                cleanupTimer.Tick += (s, e) =>
                {
                    try
                    {
                        // Close any Settings windows
                        foreach (Process proc in Process.GetProcessesByName("SystemSettings"))
                        {
                            try { proc.CloseMainWindow(); } catch { }
                            try { proc.Kill(); } catch { }
                            System.Console.WriteLine("Closed Settings window");
                        }
                        
                        // Close any Control Panel windows - this requires finding explorer windows with specific titles
                        foreach (Process proc in Process.GetProcessesByName("explorer"))
                        {
                            // Try to determine if it's a control panel window
                            try
                            {
                                IntPtr hwnd = proc.MainWindowHandle;
                                if (hwnd != IntPtr.Zero)
                                {
                                    const int nChars = 256;
                                    StringBuilder windowTitle = new StringBuilder(nChars);
                                    GetWindowText(hwnd, windowTitle, nChars);
                                    
                                    // If this is a control panel/personalization window
                                    if (windowTitle.ToString().Contains("Personalization"))
                                    {
                                        PostMessage(hwnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                                        System.Console.WriteLine("Closed Personalization window");
                                    }
                                }
                            }
                            catch { }
                        }

                        // Notification that theme was applied
                        trayIcon.ShowBalloonTip(
                            2000,
                            "Theme Applied",
                            $"Switched to {Path.GetFileNameWithoutExtension(themePath)} theme",
                            ToolTipIcon.Info);

                        // Dispose the timer
                        cleanupTimer.Stop();
                        cleanupTimer.Dispose();
                    }
                    catch (Exception ex)
                    {
                        System.Console.WriteLine($"Cleanup failed: {ex.Message}");
                    }
                };
                
                // Start the cleanup timer
                cleanupTimer.Start();
                
                System.Console.WriteLine($"Applied theme: {Path.GetFileName(themePath)}");
            }
            catch (Exception ex)
            {
                System.Console.WriteLine($"Theme application failed: {ex.Message}");
            }
        }

        // P/Invoke declarations for window handling
        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        // Windows message constants
        private const uint WM_CLOSE = 0x0010;

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