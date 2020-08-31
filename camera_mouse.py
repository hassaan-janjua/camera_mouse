# README
# Prerequisites for mouse control using camera on windows 
# Install python 
# Install pip from https://pip.pypa.io/en/stable/installing/
# pip install numpy
# install the appropriate version of opencv and opencv contrib 
#   from https://www.lfd.uci.edu/~gohlke/pythonlibs/
# pip install imutils
# pip install pywin32


import win32api, win32con
import win32gui
import numpy as np
import cv2
import imutils
import ctypes

xSpeed = 10
ySpeed = 10

# Get screen dimentions
user32 = ctypes.windll.user32
DisplayResolutionX = user32.GetSystemMetrics(78) # 1920
DisplayResolutionY = user32.GetSystemMetrics(79) # 1080

# My Yellow Highlighter
#mouseColor = (108, 169, 173)

# Red
mouseColor = (0, 0, 255)

def updateMouse(p1, p2):
    init = win32gui.GetCursorPos()
    x = init[0] + (p1[0] - p2[0]) * xSpeed
    y = init[1] + (p2[1] - p1[1]) * ySpeed

    if (x < 0):
        x = 0
    if (y < 0):
        y = 0

    if (x > DisplayResolutionX):
        x = DisplayResolutionX

    if (y > DisplayResolutionY):
        y = DisplayResolutionY

    win32api.SetCursorPos((x, y))

def mouse_driver():
  cap = cv2.VideoCapture(0)

  prev = (0,0)
  isFirst = True
  while(True):
    # Capture frame-by-frame
    ret, frame = cap.read()

    red = np.zeros(frame.shape, np.float32)

    red[:] = mouseColor

    diff_frame = red - frame

    r = diff_frame[:,:,0]
    g = diff_frame[:,:,1]
    b = diff_frame[:,:,2]

    r2 = np.square(r)
    g2 = np.square(g)
    b2 = np.square(b)

    d = np.sqrt(r2 + g2 + b2)

    ret, thresh = cv2.threshold(d, 128, 255, cv2.THRESH_BINARY_INV)
    
    thresh = thresh.astype(np.uint8)

    # erode away small noise
    kernel = np.ones((5, 5), np.uint8)
    thresh = cv2.erode(thresh, kernel, iterations=1)
    
    # dilate the pixels to connect broken components of the mouse
    kernel = np.ones((3, 3), np.uint8)
    thresh = cv2.dilate(thresh, kernel, iterations=5)
    
    # Find connected components in the threshholded image
    contours  = cv2.findContours(thresh.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    cnts = imutils.grab_contours(contours)

    if (len(cnts) > 0):
        # Assume that there is only one contour in the image
        c = cnts[0]
        
        # compute the center of the contour
        M = cv2.moments(c)
        cX = int(M["m10"] / M["m00"])
        cY = int(M["m01"] / M["m00"])

        curr = (cX, cY)
        
        if not isFirst:
          updateMouse(prev, curr)
        
        prev = (cX, cY)
        isFirst = False
    else:
        # Start over when the cursor goes out of screen
        isFirst = True

    # Display the original frame
    # Mirror the image to show movement in te right direction
    cv2.imshow('original',cv2.flip(frame, 1))
    cv2.imshow('thresh',cv2.flip(thresh, 1))

    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

  # When everything done, release the capture
  cap.release()
  cv2.destroyAllWindows()


if __name__ == "__main__":
  mouse_driver()
