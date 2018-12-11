# CenterTaskbar

![Gif](https://user-images.githubusercontent.com/3608298/49763502-d431d980-fc9a-11e8-9ac5-a7d2c52ef1c4.gif)

----
## Default Usage
Run the program and let it run in the background. It uses Windows UIAutomation to monitor for position changes and calculate a new position to center the taskbar items.

----
## Command Line Args
Specify any number as the sole command line argument to set the refresh rate in hertz. Recommended to sync to your monitor refresh rate. When no changes are being made the refresh rate drops to 10 to minimize background CPU usage. Default 60
