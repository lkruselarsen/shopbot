from pynput.mouse import Listener as ms
from pynput.keyboard import Listener as kb
from pynput.keyboard import Key, Controller
import keyboard
import os
import sys
import subprocess
import time

keyboard1 = Controller()

def on_press(key):
    while keyboard.is_pressed('q'):
        keyboard1.press(Key.f5)
        keyboard1.release(Key.f5)
        print('{0} pressed'.format(
        key))
        time.sleep(0.1)
        break
    while keyboard.is_pressed('a'):
        subprocess.call("C:/---your-path----/actualinput.vbs", shell=True)
        print('{0} pressed'.format(
        key))
        time.sleep(0.05)
        break
    while keyboard.is_pressed('z'):
        subprocess.call("C:/---your-path----/autofillscript.vbs", shell=True)
        print('{0} pressed'.format(
        key))
        time.sleep(0.05)
        break
    if key == Key.esc:
        return False
    

def on_release(key):
    print('{0} release'.format(
        key))
    if key == Key.esc:
        return False
def on_click(x, y, button, pressed):
    if pressed:
        pass
    else:
        #subprocess.call("C:/---your-path----/actualinput.vbs", shell=True)
        print('Released')

def on_scroll(x, y, dx, dy):
    print('Scrolled {0}'.format(
        (x, y)))

# Collect events until released
with ms(
        on_click=on_click,
        on_scroll=on_scroll) as listener:
    with kb(
            on_press=on_press,
            on_release=on_release) as listener:
        listener.join()
