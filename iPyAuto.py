import os
import time

import cv2
import numpy
import subprocess
import uiautomation as automation

# pip install pywin32
import win32api, win32con

path = "current_screen.png"

class iPyAuto(object):
    def screen_capture(self):
        time.sleep(2)
        c = automation.GetRootControl()
        if c.CaptureToImage(path):
            automation.Logger.WriteLine('capture image: ' + path)
        else:
            automation.Logger.WriteLine('capture failed', automation.ConsoleColor.Yellow)

    def mouse_click(self,x,y):
        win32api.SetCursorPos((x,y))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)

    def mouse_double_click(self,x, y):
        self.mouse_click(x, y)
        self.mouse_click(x, y)

    def find_object_in_screen(self,object_image, retry = 3):
        for each_try in range(retry):
            self.screen_capture()
            os.remove(path)
            image = cv2.imread(object_image)
            width, height = image.shape[:-1]

            matched_object = cv2.matchTemplate(cv2.imread(path), image, cv2.TM_CCOEFF_NORMED)
            locations = numpy.where(matched_object >= .8)
            for each_point in zip(*locations[::-1]):
                x = each_point[0] + height/2
                y = each_point[1] + width/2
                return (x, y)
            print "Could not find object(" + object_image + "). Retrying("+ str(each_try)+ ") ...."
        print "Could not find object(" + object_image + "). Terminating...."
        return False


object_name = "My Computer"
c_drive = "c drive"

# object_name = "My Computer new"
# c_drive = "c drive"

#SAMPLE

subprocess.Popen('explorer.exe')
iBuilder = automation.WindowControl(searchDepth = 1, ClassName = 'WindowsForms10.Window.8.app.0.33c0d9d')

ipyauto = iPyAuto()

object_coordinates = ipyauto.find_object_in_screen(object_name + ".png")
ipyauto.mouse_double_click(object_coordinates[0], object_coordinates[1])

object_coordinates = ipyauto.find_object_in_screen(c_drive + ".png")
ipyauto.mouse_double_click(object_coordinates[0], object_coordinates[1])
