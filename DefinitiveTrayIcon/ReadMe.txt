DEFINITIVE TRAY ICON v. 0.3
By: Phil Ciebiera
----------------------------
If you like it give me a vote,

The goal of this is to demonstrate a more effective usage of the System Tray Icon in VB.

Learn from it and enjoy.


Thanks to Mark Mokoski for explaining the right way to use balloon tips from VB,
see his code on PSC.

Thanks to LaVolpe for pointing out several issues. (:

Class Module coming soon, I hope (;

-Phil


New Features for v. 0.3
-----------------------
-Multiple Tray Icons supported, Demo Included
-WM_TRAYHOOK is now dynamically created from a generated GUID
-Created a new and improved GlobalMessageCatcher, for subclassing multiple forms at once



New Features for v. 0.2
-----------------------
-You can now have a user defined icon in a balloon tip
-You can detect when a user clicks your balloon
-You can make a balloon popup with or without sound

Bugs Squashed in v. 0.2
-----------------------
-Issue of using reserved uID, thanks LaVolpe
-Looked at a newer MSDN and changed the WM_TRAYHOOK to branch from WM_APP instead of the old WM_USER



New Features for v. 0.1
-----------------------
-Explorer.exe crash detection (tray icon persists)
-Native Windows Balloon Tips (with 4 icon styles or no icon)
-Low level detection of window messages
-Ability to use multiple tray icons for 1 form (not shown)