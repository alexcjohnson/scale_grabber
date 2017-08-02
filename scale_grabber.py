##############################################################################
#                                                                            #
# Scale grabber - grab weights from Adam CPWplus scales and type them into   #
#  another application.                                                      #
#                                                                            #
# Installation (Windows only):                                               #
# - Install Python 3.6+                                                      #
# - In a command prompt:                                                     #
#     pip install pywin32 pyHook pyserial PyAutoGUI                          #
# - Copy this file somewhere on the local machine                            #
#   (if we like this, the whole thing could be made pip installable)         #
#                                                                            #
# Use                                                                        #
# - How to specify the COM port in use?                                      #
# - Run this file (can we just double click? do we need a launcher?)         #
# - Leave the window open (minimized is fine - can we have no window?        #
#   that seems dangerous, as people will not know if the program stops       #
#   and needs restarting.)                                                   #
#                                                                            #
##############################################################################


import pyautogui
import pyHook
import argparse
import serial


def get_weight(conn, timeout=1):
    line = conn.readline()
    parts = line.split(' ')
    for part in parts:
        try:
            val = float(part)
            return val
        except ValueError:
            pass

    if line and line.strip():
        print('Data received but no weight found:')
        print(line)

    return None  # we didn't find a value (timeout?)


def type_weight(weight):
    pyautogui.typewrite('{:.2f}'.format(weight))


def grab_weight(conn):
    weight = get_weight(conn)
    if weight is not None:
        type_weight(weight)


def request_weight(conn):
    # clear any previous waiting data
    conn.reset_input_buffer()

    # ask for gross weight (should it be N for net instead?)
    conn.write('G\r\n')


def keypress_handler(conn, key):
    def on_key_event(event):
        if event.Key == key:
            grab_weight(conn)
        # regardless, let the event continue propagating
        return True

    return on_key_event


def watch_keyboard(conn, key):
    # on pressing a special key (f12?) you can request a weight
    # this is an alternative to pressing "print" on the scale itself
    # is this easier?
    hm = pyHook.HookManager()
    hm.KeyDown = keypress_handler(conn, key)
    hm.HookKeyboard()

    # TODO: needed? What does this do? Is this the infinite event loop?
    import pythoncom
    pythoncom.PumpMessages()


# blocking - if we want to watch both keyboard and line, do we need
# to put this in HookManager somehow as a periodic callback?
def watch_line(conn):
    while True:
        grab_weight(conn)


if __name__ == '__main__':
    desc = ('Grab weights from Adam CPWplus scales and type them into '
            'another application')
    parser = argparse.ArgumentParser(description=desc)
    parser.add_argument('port', help=('the port the scale is connected to, '
                                      'eg: COM3'))
    parser.add_argument('--key', help=('if provided, take the weight when '
                                       'this key is pressed, rather than when '
                                       'PRINT is pressed on the scale itself. '
                                       'eg: f12'))

    args = parser.parse_args()

    port = args.port

    with serial.Serial(port) as conn:
        if args.key:
            watch_keyboard(conn, args.key)
        else:
            watch_line(conn)
