## GhostBusterPlus
A utility to improve the eInk experience on the Lenovo ThinkBook Plus Gen 4 laptop.

### NEW: The program has been completely rewritten in version 0.2 to use an entirely new approach!

![logo](https://github.com/user-attachments/assets/def07e61-7fb3-45f7-be4b-e5d81a81018c)

## Features (Version 0.3)
- Automatically activate the clear ghosts keyboard shortcut (e.g. F4), to refresh the eInk display.
  - The change detection is now done by a DirectX shader running on the GPU that stores and compares changes in the screen. If the modified pixels exceed some threshold (e.g. 3%), a refresh is queued.
  - Version 0.3 fixed a bug whereby the DirectX context was not restarted upon a display switch, causing silent failures in change detection.
  - The screen information is never transfered into the program, but rather all screen bitmap processing occurs on the GPU in DirectX, and the program only receives the percent change. The app also has no communications or internet connectivity and does not require elevated privileges.
- The screen refresh is delayed while the user is actively doing something, such as providing keyboard input or (optionally) moving the mouse cursor. Also, scrolling or rapidly updating screens will delay refresh in order to avoid a rapid refresh condition.
- Many options for customizing the screen change threshold, input/screen refresh delay time, etc.
- A menu option and hotkey for restarting the EInkApp, which tends to sometimes die or exhibit weird behavor when switching screens.
- Ability to quickly switch between the provided eInk (light high contrast) and Dark Windows 11 themes.
- Automatic detection of the active display (eInk or OLED) and automatic theme switching

## Installation
First, you need to have the latest version of EInkPlus installed on your system. The latest version provides shortcut keys for various functions, such as clearing ghosts. Open the EInkApp Settings -> Display Settings -> Shortcut Key Settigns -> Edit (button). Change "Clear Eink Ghosts" to the "F4" key.
 
 Double click on GhostBusterPlus.exe. The app should launch a taskbar tray icon, which might be hidden in the overflow area unless you drag it onto the main taskbar.
 The icon looks like the Ghostbusters logo. You can right click (two finger click) on the icon to pull up a menu listing the options, configurations and hotkeys, or to quit.

 ### Run automatically at startup
 To have Windows launch GhostBusterPlus automatically when you log in, do the following:
 - Open the folder where you extracted the GhostBusterPlus zip.
 - Press Win+R to bring up the Run dialog box. Enter "shell:startup" and hit OK.
 - Drag GhostBusterPlus.exe into the Startup folder while holding down the Alt key to create a shortcut.

 ## To-do
 - None planned at this time. Automatic display detection and theme switching was implemented in v0.3.

## System Requirements
- Lenovo ThinkBook Plus Gen 4
  - Should work with the Gen 2 with very minimal changes, possibly.
- Windows 11
  - Tested on update 24H2 with eInkPlus app version 1.0.124.3

## Bugs and Feature Requests
- Open up an Issue in this GitHub repository to report issues or bugs.

## Security and Warnings
This app has no internet access, communication, location or recording functionality. The screen change detection is performed entirely in DirectX on the GPU and the images are never transfered to the program (CPU). Instead, a DirectX shader computes the percent change and returns that number to the program. 

If Microsoft Edge, Windows Defender or other AV program gives you trouble, it’s probably because it’s a new app and it is not signed (requires paying $$$ to Microsoft).
You are free to inspect the source code and compile the project in Visual Studio 2022 yourself. That said, I'm not a professional Windows developer, so I make no guarantees,
provide no warranties or any other promises of performance or fitness.

## Latest Version of EInkPlus
For some reason, Lenovo is not keeping the drivers & support website up-to-date with the latest version of EInkPlus. This may be because it is supposed to automatically update itself via an OTA update. However, this often does not happen correctly, and the latest version is required for compatibility with Windows 11 24H2. As mentioned, the latest version also adds the shortcut key feature (e.g. F4), which is required for GhostBusterPlus >=0.2 to function. 

Fortunately, the latest version, LenovoEinkPlus_OTA2_PRC_1.0.124.3, of EInkPlus can be found either [here](https://drive.google.com/file/d/117gDwTUzBHfVHCzwmdNqCLCmyMuzL4Ps/view?usp=sharing) or [here](https://forums.lenovo.com/t5/ThinkBook-Plus-Laptops/EInk-Plus-Driver-Reader-Note-Apps-for-ThinkBook-Plus-Gen-4-Win11-24H2-Compatible/m-p/5377868).

## License
Copyright (c) 2025 by the author, joncox123. All rights reserved.
