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
from time import sleep
from functools import partial
import pygame
import sys

class KeyManager:
    def __init__(self):
        self.keys_down = set()
        self.latest_down = None
        self.watched_key_str = None
        self.callback = None
        self.quit_key_str = 'control-c-x'

    def clean_key_name(self, event):
        # strip l/r off shift/control
        key = event.Key.lower()
        if len(key) > 1 and key[0] in ('l', 'r'):
            key = key[1:]
        return key

    def mark_key_down(self, event):
        key = self.clean_key_name(event)
        self.keys_down.add(key)
        self.latest_down = key
        if self.watched_key_str and self.callback:
            if self.keys_match(self.watched_key_str):
                self.callback()
                return 0  # stop this event propagating

        if self.keys_match(self.quit_key_str):
            sys.exit(1)

        return 1  # any other key situation - let it propagate

    def mark_key_up(self, event):
        self.keys_down.discard(self.clean_key_name(event))
        self.latest_down = None
        return 1  # always let this event continue propagating

    def keys_match(self, key_str):
        key_combo = key_str.split('-')
        key_combo_set = set(key_combo)
        return (self.latest_down in key_combo and
                (len(self.latest_down) == 1 or  # so no control or shift etc
                 self.latest_down == key_combo[-1]) and
                key_combo_set.issubset(self.keys_down))

    def watch_for_key(self, hook_manager, key_str, callback):
        self.watched_key_str = key_str.lower()
        self.callback = callback
        hook_manager.KeyDown = self.mark_key_down
        hook_manager.KeyUp = self.mark_key_up
        hook_manager.HookKeyboard()

def get_weight(conn):
    # wait for data to arrive from a request
    sleep(0.1)
    lines = str(conn.read_all()).strip()
    parts = lines.split(' ')
    for part in parts:
        try:
            val = float(part)
            return val
        except ValueError:
            pass

    if lines:
        print('Data received but no weight found:')
        print(lines)
    else:
        print('no data received')

    return None  # we didn't find a value


def type_weight(weight):
    print('~{:.1f}~'.format(weight))
    pyautogui.typewrite('{:.1f}'.format(weight))


def grab_weight(conn):
    weight = get_weight(conn)
    if weight is not None:
        type_weight(weight)


def request_weight(conn):
    # clear any previous waiting data
    conn.reset_input_buffer()

    # ask for gross weight (should it be N for net instead?)
    conn.write(b'G\r\n')

    # then grab the result
    grab_weight(conn)


def watch_keyboard(conn, key):
    # on pressing a special key (f12?) you can request a weight
    # this is an alternative to pressing "print" on the scale itself
    # is this easier?
    hm = pyHook.HookManager()
    km = KeyManager()

    km.watch_for_key(hm, key, partial(request_weight, conn))

    pygame.init()
    while True:
        sleep(0.05)
        pygame.event.pump()


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
                                       'eg: control-add'))

    args = parser.parse_args()

    port = args.port

    with serial.Serial(port) as conn:
        print('Leave this program running, but you can minimize it')
        if args.key:
            print('press ' + args.key + ' to insert the current weight at your cursor')
            print('press control-c-x all together to quit')
            watch_keyboard(conn, args.key)
        else:
            print('press the print button on the scale to insert the current weight at your cursor')
            print('press ctrl-c to quit')
            watch_line(conn)

# C:\Users\Volunteer>c:\Users\Volunteer\AppData\Local\Programs\Python\Python36-32\python.exe c:\scale_grabber\scale_grabber.py COM3 --key control-add