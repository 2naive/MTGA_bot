# Run as admin
# Make console window pinned and transparent with Wintop
#
# Tested for screen resolution of 1920 x 1080
#
# You have to be in main menu with deck and game type selected before run
# You may see also:
#   https://github.com/defaultroot1/Python-Scripts/blob/master/mtga_bot.py
#   https://defaultroot.com/index.php/2020/07/07/a-magic-the-gathering-arena-bot-in-python/
#   https://github.com/Lyoshikawa3914/MTGA_bot_public
#   https://github.com/hornsilk/BreyaBot

import os
import sys
import time
import win32api, win32con, win32gui
from win32api import GetSystemMetrics

import numpy
from numpy import *

# Split multiple delimiters
import re

import pynput
from pynput.keyboard import Key, Controller
from pynput.mouse import Controller, Listener, Button

# Screenshots and image operations
from PIL import Image, ImageGrab, ImageOps

# Optical Character Recognition (OCR)
import pytesseract
import cv2

# https://stackoverflow.com/questions/44398075/can-dpi-scaling-be-enabled-disabled-programmatically-on-a-per-session-basis/44422362#44422362
import ctypes
# Set DPI Awareness  (Windows 10 and 8)
ctypes.windll.shcore.SetProcessDpiAwareness(2)
# Set DPI Awareness  (Windows 7 and Vista)
ctypes.windll.user32.SetProcessDPIAware(2)

WINDOW_NAME = "MTGA"
SCREEN_WIDTH = GetSystemMetrics(0) # 3840
SCREEN_HEIGHT = GetSystemMetrics(1) # 2160
ENDGAME_COUNTER_TRESHOLD = 3
COLOR_SENSITIVITY = 25
CARD_WIDTH = math.ceil(SCREEN_WIDTH * 0.08984375) # avg(range(325,365)) = 345 (342.5)
TICK = 3

class Cords():
    home_button         = (math.ceil(SCREEN_WIDTH * 0.092188),  math.ceil(SCREEN_HEIGHT * 0.03426))
    play_button         = (math.ceil(SCREEN_WIDTH * 0.895),     math.ceil(SCREEN_HEIGHT * 0.938))
    attack_button       = (math.ceil(SCREEN_WIDTH * 0.8815),    math.ceil(SCREEN_HEIGHT * 0.86666)) # should be lower part
    enemy_icon          = (math.ceil(SCREEN_WIDTH / 2),         math.ceil(SCREEN_HEIGHT * 0.1))
    loose_3_life_button = (math.ceil(SCREEN_WIDTH / 2),         math.ceil(SCREEN_HEIGHT * 0.23148))
    battle_middle_card  = (math.ceil(SCREEN_WIDTH * 0.47917),    math.ceil(SCREEN_HEIGHT * 0.5463)) # discard battle card
    artifact            = (math.ceil(SCREEN_WIDTH * 0.71875),    math.ceil(SCREEN_HEIGHT * 0.69444)) # 2760x1500
    artifact_activate   = (math.ceil(SCREEN_WIDTH * 0.71875),    math.ceil(SCREEN_HEIGHT * 0.6412)) # 2760x1385

def getColorDistance(color1, color2):
    return sum(abs(color1[0]-color2[0])+abs(color1[1]-color2[1])+abs(color1[2]-color2[2]))

def isColorEqual(color1, color2):
    #print('C{0}'.format(sum(color1)-sum(color2)), end='')
    return color1 == color2 or getColorDistance(color1, color2) < COLOR_SENSITIVITY

def win32getColor(coordinates):
    hWin = win32gui.GetActiveWindow() # GetDesktopWindow
    # hWinDC = win32gui.GetWindowDC(hwn)
    # https://www.equestionanswers.com/vcpp/getdc-getwindowdc.php
    # https://www.programcreek.com/python/example/89843/win32gui.ReleaseDC
    hWinDC = win32gui.GetDC(hWin)
    color = rgbint2rgbtuple(win32gui.GetPixel(hWinDC, coordinates[0] , coordinates[1]))
    win32gui.ReleaseDC(hWin, hWinDC)
    return color

def rgbint2rgbtuple(RGBint):
    red =  RGBint & 255
    green = (RGBint >> 8) & 255
    blue =   (RGBint >> 16) & 255
    return (red, green, blue)

def getMousePos():
    x, y = win32api.GetCursorPos()
    print(x, y)
    return (x, y)

def makeScreenshot():
    # takes a full snapshot of the screen
    im = ImageGrab.grab()
    im.save(os.getcwd() + '\\screenshot_' + time.strftime("%Y-%m-%d_%H-%M") + '.png', 'PNG')

def leftClick():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    time.sleep(.1)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    time.sleep(.1)
    #print('Click at {0}'.format(win32api.GetCursorPos()))

def doubleClick():
    leftClick()
    leftClick()

def press(button):
    kb = pynput.keyboard.Controller()
    kb.press(button)
    time.sleep(.1)
    kb.release(button)
    time.sleep(.1)

def leftDown():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    time.sleep(.1)

def leftUp():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    time.sleep(.1)

def mousePos(cord):
    win32api.SetCursorPos((cord[0], cord[1]))
    # ctypes.windll.user32.SetCursorPos(cord[0], cord[1])
    time.sleep(.1)
    # print('Move to {0}'.format((cord[0], cord[1])))

def mouseScroll():
    mouse = Controller()
    mouse.scroll(10, 0)

def image2text(image):
    time.sleep(6)
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

    im = ImageGrab.grab(image)
    #resize_im = im.resize((830, 55), resample=Image.NEAREST) #(width, height)

    text = pytesseract.image_to_string(cv2.cvtColor(numpy.array(im), cv2.IMREAD_GRAYSCALE), lang='eng')

    #im.show()
    print(text)
    split_text = text.split()
    print(split_text)
    return split_text

class bot():
    def __init__(self, email = None, password = None):
        self.email = email
        self.password = password
        self.endGameCounter = 0
        self.endGameTreshold = ENDGAME_COUNTER_TRESHOLD
        self.timeGameStarted = 0
        self.timeGameTreshold = 1200
        self.timeWaitingStarted = 0
        self.timeWaitingTreshold = 120

    def run(self):
        while True:
            self.start()
            self.loading()
            #self.loading_screen()
            #self.keep_cards()
            self.play()

    def start(self):
        print(time.strftime("%Y-%m-%d %H:%M"), 'Starting')
        # if not in main screen
        time.sleep(1)
        mousePos(Cords.home_button)
        leftClick()
        time.sleep(TICK)
        # Protecting window from backgrounding
        # https://stackoverflow.com/questions/2090464/python-window-activation
        # https://stackoverflow.com/questions/56857560/win32gui-setforegroundwindowhandle-not-working-in-loop
        win32gui.SetForegroundWindow(win32gui.FindWindow(None, WINDOW_NAME))
        win32gui.SetActiveWindow(win32gui.FindWindow(None, WINDOW_NAME))
        mousePos(Cords.play_button)
        # in case of additional popup
        time.sleep(0.5)
        leftClick()
        press(Key.space)
        # end
        # in case claiming prize
        time.sleep(0.5)
        leftClick()
        press(Key.space)
        # end
        time.sleep(0.5)
        leftClick()
        time.sleep(0.5)
        leftClick()
        return

    def loading(self):
        self.timeWaitingStarted = time.time()
        time.sleep(3)
        print("Waiting to play ", end='', flush=True)
        while not self.checkButtonColor() == 'red':
            press(Key.space)
            time.sleep(TICK)
            print('.', end='', flush=True)
            if time.time() - self.timeWaitingStarted > self.timeWaitingTreshold:
                print(time.strftime("%Y-%m-%d %H:%M"), 'Time passed. Exiting ...', end='', flush=True)
                break
            #print('Loading color:', reference_color)
        #print("Loading new color:", color)
        print(" Ready")
        time.sleep(1)
        return

    def play(self):
        self.timeGameStarted = time.time()
        print("Playing ", end='')
        reference_color = win32getColor((0, 0))
        #print('Play reference color:', reference_color)
        while self.isGameRunning(reference_color):
            if self.checkButtonColor() == 'black':
                print('.', end='', flush=True)
                continue
            elif self.checkButtonColor() == 'blue':
                print('R', end='', flush=True)
                press(Key.space)
                # in case it's two choice button and not clickable by space
                mousePos(Cords.attack_button)
                leftClick()
            elif self.checkButtonColor() == 'red':
                time_start = time.time()
                print('A', end='', flush=True)
                self.clickAllCards()
                self.clickAllCards()
                # Activate vortex
                mousePos(Cords.artifact)
                leftClick()
                mousePos(Cords.artifact_activate)
                leftClick()
                time.sleep(TICK)
                # End Activate artifact
                press(Key.space)
                # if placewalker -> click enemy icon
                self.clickEnemy()
                press(Key.space)
                print(math.ceil(time.time() - time_start), end='', flush=True)
        print()
        print(time.strftime("%Y-%m-%d %H:%M"), 'Endgame in {0}m'.format(self.getTimePassed()), flush=True)
        self.endGameCounter=0
        mousePos((3, 3))
        leftClick()
        time.sleep(15)
        leftClick()
        time.sleep(2)
        leftClick()
        time.sleep(1)

    def isGameRunning(self, reference_color):
        if time.time() - self.timeGameStarted > self.timeGameTreshold:
            print(time.strftime("%Y-%m-%d %H:%M"), 'Time passed. Exiting ...', end='', flush=True)
            return False
        time.sleep(TICK)
        #print('Game reference color:', reference_color, sum(reference_color))
        color = win32getColor((0, 0))
        #print('Play color:', color, sum(color))
        #print('.', end='')
        if isColorEqual(color, reference_color):
            self.endGameCounter=0
            return True
        else:
            self.endGameCounter+=1
            print('E{0}'.format(self.endGameCounter), end='', flush=True)
            # if first time -> just count and skip

            if self.endGameCounter == 1:
                return True

            if self.endGameCounter >= 2:
                makeScreenshot()

            # in case it's popup -> close it
            if not self.checkButtonColor() == 'red' and not self.checkStartButtonColor() == 'red':
                press(Key.space)

            # if Discard card -> click middle card and press submit
            if not self.checkButtonColor() == 'red' and not self.checkStartButtonColor() == 'red':
                mousePos(Cords.battle_middle_card)
                leftClick()
                press(Key.space)

            # if Discard hand card -> click middle hand card and press submit
            if not self.checkButtonColor() == 'red' and not self.checkStartButtonColor() == 'red':
                mousePos((math.ceil(SCREEN_WIDTH / 2), SCREEN_HEIGHT - 1))
                leftClick()
                press(Key.space)

            # if Discard 2 hand card -> click second middle hand card and press submit
            if not self.checkButtonColor() == 'red' and not self.checkStartButtonColor() == 'red':
                mousePos((math.ceil(SCREEN_WIDTH / 2) - CARD_WIDTH, SCREEN_HEIGHT - 1))
                leftClick()
                press(Key.space)

            # in case game was stopped by Choose one ... loose 3 life popup
            if not self.checkButtonColor() == 'red' and not self.checkStartButtonColor() == 'red':
                mousePos(Cords.loose_3_life_button)
                leftClick()

            if self.endGameCounter == self.endGameTreshold and color[0] > 250 - COLOR_SENSITIVITY:
                print('timeout?', color, end='', flush=True)

            return self.endGameCounter < self.endGameTreshold

    def getTimePassed(self):
        return math.ceil((time.time() - self.timeGameStarted) / 60)

    def checkButtonColor(self):
        color = win32getColor(Cords.attack_button)
        #print('Attack button color:', color, Cords.attack_button)
        if color[0] > 200:
            return 'red'
        elif color[2] > 200:
            return 'blue'
        else:
            return 'black'

    def checkStartButtonColor(self):
        color = win32getColor(Cords.attack_button)
        #print('Attack button color:', color, Cords.attack_button)
        if color[0] > 200:
            return 'red'
        elif color[2] > 200:
            return 'blue'
        else:
            return 'black'

    def clickAllCards(self):
        time.sleep(.1)
        cards_max = 8
        kb = pynput.keyboard.Controller()
        start_point = math.ceil(SCREEN_WIDTH / 2 + CARD_WIDTH / 4 + 3 * CARD_WIDTH) # @todo use cards_max
        cards_n = 0
        for x in range(cards_max):
            card_coordinates = (start_point - x * CARD_WIDTH, math.ceil(SCREEN_HEIGHT * 0.97))
            if not self.isCard(card_coordinates):
                continue
            cards_n+=1
            doubleClick()
            mousePos((SCREEN_WIDTH-1, SCREEN_HEIGHT-1))
            time.sleep(.5)
            press('z')
            leftClick()
            time.sleep(.15)
        print('[{0}]'.format(cards_n), end='', flush=True)

    def isCard(self, card_coordinates):
        #time.sleep(2)
        color1 = win32getColor(card_coordinates)
        mousePos(card_coordinates)
        color2 = win32getColor(card_coordinates)
        d = getColorDistance(color1, color2)
        #print(color1, color2, d, not isColorEqual(color1, color2), flush=True)
        return not isColorEqual(color1, color2)

    def clickEnemy(self):
        mousePos(Cords.enemy_icon)
        leftClick()

# https://pythonhosted.org/pynput/mouse.html
def on_move(x, y):
    print('Pointer moved to {0}'.format((x, y)))

def on_click(x, y, button, pressed):
    if button==Button.right:
        if pressed:
            time.sleep(.1)
            im = ImageGrab.grab()
            color = im.getpixel((x, y))

            x2, y2 = win32api.GetCursorPos()
            color2 = win32getColor((x2, y2))

            print('{0} at {1} color {2}; win32api coordinates {3} color {4}'.format(button, (x, y), color, (x2, y2), color2))
        return True
    else:
        return False

def on_scroll(x, y, dx, dy):
    print('Scrolled {0}'.format((x, y)))

# Collect events until released

if (len(sys.argv) > 1 and sys.argv[1]):
    print('RightClick to get coordinates. LeftClick for exit')
    with Listener(
    #        on_move=on_move,
    #        on_scroll=on_scroll,z
            on_click=on_click) as listener:
        listener.join()
else:
    bot().run()
