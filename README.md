Scale grabber - grab weights from Adam CPWplus scales and type them into another application.

Installation (Windows only):
- Install Python 3.6+
- In a command prompt:
```
pip install pywin32 pyHook pyserial PyAutoGUI
```
- Copy this file somewhere on the local machine
  (if we like this, the whole thing could be made pip installable)

Use
- How to specify the COM port in use?
- Run this file (can we just double click? do we need a launcher?)
- In a command prompt, navigate to the directory where you put this file
- Type (assuming your scale is connected to COM3):
```
python scale_grabber.py COM3
```
- Optionally, to trigger weighing from the keyboard (f12 for example) rather than from the scale's built-in PRINT button:
```
python scale_grabber.py COM3 --key f12
```
- Leave the window open (minimized is fine - can we have no window? That seems dangerous, as people will not know if the program stops and needs restarting.)
