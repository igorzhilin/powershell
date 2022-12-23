# adb helpers
## What is adb?
adb (Android debug bridge) is a tool to manage Android devices from a PC using a command line. Hence, from PowerShell, too! 👍

`adb.exe` is part of so-called Android SDK Platform-Tools package - [download it from here](https://developer.android.com/studio/releases/platform-tools).

## What can you do with adb?
Plenty of things, but for me mainly it is an easy way to use PC command line to:
* [mass install apks (Android applications) from PC](adb%20mass%20install%20apks.ps1)
* [mass uninstall apps from the Android device](adb%20mass%20uninstall.ps1) (e.g. preinstalled vendor bloatware, e.g. from Samsung, Lenovo, Huawei, Facebook, etc.)
* [mass copy ("push") files and directories to Android device](adb%20mass%20push.ps1)

## How to set up your Android device to support adb?
1. Enable Developer Options. This is done by tapping on Android Build Number 7 times. Typically you find the **Build Number** in the **Settings > About > System**. Tap on it 7 times. Then an "Enter PIN" dialog will appear (if you have device PIN activated), and when you complete it, there will be a message (toast) "You are now a developer". Then you will find **Developer Options** in **Settings** or **Settings > System**.
2. In the **Developer Options**, activate **USB Debugging**. This enables the support for adb on the Android device.
3. In the same **Developer Options**, disable security check for apps installed over ADB. Some devices (e.g. Huawei MatePad 10.4) have this option. If you do not disable this check, then every time you run `adb install`, the device will show a confirmation window where you have to tap **Install** - not ideal for unattended installations. 
4. Finally, allow the device to talk to the PC over adb. Plug the USB cable in the Android device and the PC. Run command `adb devices` on the PC. A prompt will appear on the Android device: **Trust this PC?** Say Yes, and you are good to go.