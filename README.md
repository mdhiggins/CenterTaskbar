# CenterTaskbar

![Gif](https://user-images.githubusercontent.com/3608298/49901443-36234800-fe2f-11e8-89dd-9ab609a34fba.gif)

----
## Default Usage
Run the program and let it run in the background. It uses Windows UIAutomation to monitor for position changes and calculate a new position to center the taskbar items.

## Command Line Args
Specify any number as the sole command line argument to set the refresh rate in hertz. Recommended to sync to your monitor refresh rate. When no changes are being made the refresh rate drops to 10 to minimize background CPU usage. Default 60
