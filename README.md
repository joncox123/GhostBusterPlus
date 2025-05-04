## GhostBusterPlus
A utility to improve the eInk experience on the Lenovo ThinkBook Plus Gen 4 laptop.

![logo](https://github.com/user-attachments/assets/def07e61-7fb3-45f7-be4b-e5d81a81018c)

## Features
- Automatically activate the floating button in the EInkPlus app, to refresh the eInk display, under the following conditions:
   - A window is resized or moved
   - The foreground window changes
   - User scrolls inside a window (works for many but not all apps. Not yet working for Chrome, Edge or Windows Explorer. Works for NotePad and Visual Studio 2022)
- Manually refresh screen with Ctrl+Shift+D
- Toggle the automatic refresh on/off with Ctrl+Shift+Z
- Restart the EInkPlus app in the event it malfunctions with Ctrl+Shift+X
- Do not refresh the screen while the user is "doing something", such as typing, scrolling or moving windows.
   - Set and save the timeout for automatic screen refresh. The default is 1500 ms, configurable from 1 second to 20 seconds.

 ## Installation
 Double click on GhostBusterPlus.exe. The app should launch a taskbar tray icon, which might be hidden in the overflow area unless you drag it onto the main taskbar.
 The icon looks like the Ghostbusters logo. You can right click (two finger click) on the icon to pull up a menu listing the options, configurations and hotkeys, or to quit.

 ### Run automatically at startup
 To have Windows launch GhostBusterPlus automatically when you log in, do the following:
 - Open the folder where you extracted the GhostBusterPlus zip.
 - Press Win+R to bring up the Run dialog box. Enter "shell:startup" and hit OK.
 - Drag GhostBusterPlus.exe into the Startup folder while holding down the Alt key to create a shortcut.

 ## To-do
 - Add a hot-key and menu option for toggling between an eInk friendly high contrast Windows theme and a dark theme
    - Works in AutoHotKey v2, but I need to figure out how to get this to work properly with C#
 - Figure out how to activate he floating refresh button without quickly moving the cursor, clicking the button and moving it back. I investigated this a bit, without success, but it seems like it should be possible with UI Automation features in C#. Since the "button" is actually a clickable region in a transparent window that covers the foreground of your entire screen, its a non-trivial process.

## System Requirements
- Lenovo ThinkPad Plus Gen 4
  - Should work with the Gen 2 with very minimal changes, if any. Probably just need to change the name of the floating layover app from "4" to "2".
- Windows 11
  - Tested on update 24H2 with eInkPlus app version 1.0.124.3

## Bugs and Feature Requests
- Open up an Issue in this GitHub repository. I will not respond to e-mails, DMs or requests made in any other way. Thanks for your understanding.

## Security and Warnings
This app has no internet access, communication, location or recording (other than timeout value in ms) features. It does not require elevated permissions.
If Microsoft Edge, Windows Defender or other AV program gives you trouble, it’s probably because it’s a new app and it is not signed (requires paying $$$ to Microsoft).
You are free to inspect the source code and compile the project in Visual Studio 2022 yourself. That said, I'm not a professional Windows developer, so I make no guarantees,
provide no warranties or any other promises of performance or fitness. 

## License
Copyright (c) 2005 by the author, joncox123. All rights reserved.
