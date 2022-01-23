# CenterTaskbar

![Gif](https://user-images.githubusercontent.com/3608298/49901443-36234800-fe2f-11e8-89dd-9ab609a34fba.gif)

----
## Archived
* Windows 11 will not be supported. Underlying changes to the taskbar break the technique this program uses to center the icons so this solution no longer functions and since Windows 11 implements native centering there will be no updates to support Windows 11. I've swithced to using StartAllBack which offers way more features and cleaner animations. Archiving this project for record purposes and it will remain available to use

## Features
* Dynamic - works regardless of number of icons, DPI scaling grouping, size. All padding is calculated
* Animated - resizes along with default windows animations
* Performant - sleeps when no resizing taking place to 0% CPU usage
* Multimonitor suppport
* Vertical orientation support
* Multiple DPI support

## Usage
Run the program and let it run in the background. It uses Windows UIAutomation to monitor for position changes and calculate a new position to center the taskbar items.

## Command Line Args
First command line argument sets the refresh rate in hertz during active icon changes. Default `60`. Recommended to sync to your monitor refresh rate or higher. When no changes are being made program goes to sleep and awaits for events triggered by UIAutomation to restart the repositioning thread allowing it to drop to 0% CPU usage.

Specifically it will monitor for:
* `WindowOpenedEvent`
* `WindowClosedEvent`
* `AutomationPropertyChangedEvent: BoundingRectangleProperty`
* `StructureChangedEvent`
