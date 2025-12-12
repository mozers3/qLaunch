## ps Quick Launch
Another variant to create a Quick Launch Toolbar in Windows 11

![main](https://github.com/mozers3/qLaunch/wiki/images/main.png)

### Features:
* Works on modern Windows versions (tested on Windows 10/11)
* Can be used as a portable application
* Creates a clear cascading menu with separators
* Can be triggered either by clicking the taskbar icon or via a hotkey combo (default: `Ctrl+Alt+Q`)
* Two easy methods for adding any programs or documents to the **ps Quick Launch** menu:
    * Via a dedicated context menu (_right-click any main menu entry_)
    * Via the item "Send to" file's system context menu
* Automatic menu generation from system folders like **Quick Launch**, **Start Menu\Programs**, or any custom folder (make_JSON.ps1)
* Edit any menu item (_right-click on main menu entry_) or the entire menu (via "Edit JSON")
* `qLaunch.json` can be edited in a built-in editor with syntax highlighting, validation, and formatting
* Ability to add custom commands with parameters (e.g., hidden mode launch)
* The path to the new menu item file (or icon file) can use environment variables or a relative path (relative to the qLaunch directory). If the executable file is accessible via %PATH%, specifying the path is not necessary.
* Supports running programs as Administrator (**Elevated Mode**)

### How to Use:
1. To launch, you need the files `qLaunch.exe` and `qLaunch.json`.
2. Right-click menu items to:
    * Modify/delete existing items
    * Insert new items or separators above selection
4. To bulk add items from system folders (**Quick Launch**, **Start Menu\Programs**) or any custom folder, run in `Terminal` Windows:

![cmd](https://github.com/mozers3/qLaunch/wiki/images/cmd.png)

6. Add any file to the menu by right-clicking it, selecting "Send to" in the system context menu (_in Windows 11, hold `Shift` to reveal it_), then choosing `ps Quick Launch`.
7. Full menu customization is available via "Edit JSON" option:
    * By default, editing is done in the built-in editor
    * You can specify any external editor (e.g. "`notepad`") by setting the value of the `Editor` key in the `Settings` section of `qLaunch.json`
    * The editor will not allow to save a file damaged by editing and will point out the error you made

![editor](https://github.com/mozers3/qLaunch/wiki/images/editor.png)

8. Hotkey customization: Modify the default `Ctrl+Alt+Q` shortcut by editing the `HotKeys` value in the `Settings` section of `qLaunch.json` (_requires program restart_)
9. Hold `Shift` while clicking a menu item to run it as Administrator (**Elevated Mode**).

### Acknowledgments:
* Special thanks to [chrizonix](https://github.com/chrizonix/QuickLaunch) for the idea and implementation, which almost met my needs.
