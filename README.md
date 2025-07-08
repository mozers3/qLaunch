## ps Quick Launch
Another variant to create a Quick Launch Toolbar in Windows 11

![main](https://github.com/mozers3/qLaunch/wiki/images/main.png)

### Features:
* Works on modern Windows versions (tested on Windows 10/11)
* Creates a clear cascading menu with separators
* Can be triggered either by clicking the taskbar icon or via a hotkey combo (default: `Ctrl+Alt+Q`)
* Two easy methods for adding any programs or documents to the **ps Quick Launch** menu:
    * Via a dedicated context menu (_right-click any main menu entry_)
    * Via the item "Send to" file's system context menu
* Automatic menu generation from system folders like **Quick Launch**, **Start Menu\Programs**, or any custom folder
* Edit any menu item (_right-click on main menu entry_) or the entire menu (via "Edit Items")
* Ability to add custom commands with parameters (e.g., hidden mode launch)
* Supports running programs as Administrator (**Elevated Mode**)

### How to Use:
1. To launch, you need the files `qLaunch.exe` and `qLaunch.json`.
2. Right-click menu items to:
    * Modify/delete existing items
    * Insert new items or separators above selection
4. To bulk add items from system folders (**Quick Launch**, **Start Menu\Programs**) or any custom folder, run in CMD:
![cmd](https://github.com/mozers3/qLaunch/wiki/images/cmd.png)
5. Add any file to the menu by right-clicking it, selecting "Send to" in the system context menu (_in Windows 11, hold `Shift` to reveal it_), then choosing `ps Quick Launch`.
6. Full menu customization is available via "Edit Items" option:
    * If user-made edits to qLaunch.json contain errors, the program automatically restores the last working version
    * To validate JSON syntax after manual editing, use [JSONLint](https://jsonlint.com)
7. Change the hotkey combo (default: `Ctrl+Alt+Q`) by modifying the registry value `HotKeys` at: `HKEY_CURRENT_USER\Software\qLaunch`
8. Hold `Shift` while clicking a menu item to run it as Administrator (**Elevated Mode**).

### Acknowledgments:
* Special thanks to [chrizonix](https://github.com/chrizonix/QuickLaunch) for the idea and implementation, which almost met my needs.
