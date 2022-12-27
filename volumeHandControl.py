import cv2 as cv
import time
import numpy as np
import handTracking as htm
import math
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

cap = cv.VideoCapture(0)
cap.set(3,640)
cap.set(4,480)
ptime =0
detector = htm.handDetector(detectionCon=0.7)

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
#volume.GetMute()
#volume.GetMasterVolumeLevel()
volRange=volume.GetVolumeRange()
minVol = volRange[0]
maxVol = volRange[1]
while True:
    success, img = cap.read()
    img= detector.findHands(img)
    lmlist = detector.findPosition(img)
    if len(lmlist)!=0:
        x1,y1 = lmlist[4][1],lmlist[4][2]
        x2,y2 = lmlist[8][1],lmlist[8][2]
        cx,cy = (x1+x2)//2, (y1+y2)//2
        cv.circle(img,(x1,y1),15,(255,0,255),cv.FILLED)
        cv.circle(img,(x2,y2),15,(255,0,255),cv.FILLED)
        cv.line(img,(x1,y1),(x2,y2),(255,0,0),3)
        cv.circle(img,(cx,cy),3,(255,0,0),3,cv.FILLED)
        length = math.hypot(x2-x1,y2-y1)
        #print(length)
        vol=np.interp(length,[10,150],[minVol,maxVol])
        #print(int(length),vol)
        volume.SetMasterVolumeLevel(vol, None)
        volBar = np.interp(length, [10, 150], [400, 150])
        volPer = np.interp(length, [10, 150], [0, 100])
        if length<50:
            cv.circle(img,(cx,cy),15,(0,255,0),cv.FILLED)
            
        cv.rectangle(img, (10, 150), (85, 400), (255, 0, 0), 3)
        cv.rectangle(img, (10, int(volBar)), (85, 400), (255, 0, 0), cv.FILLED)
        cv.putText(img, f'{int(volPer)} %', (40, 450), cv.FONT_HERSHEY_COMPLEX,
                1, (255, 0, 0), 3)
    ctime = time.time()
    fps = 1/(ctime-ptime)
    ptime = ctime
    cv.putText(img,f'FPS: {int(fps)}',(30,50),cv.FONT_HERSHEY_SIMPLEX,1,(233,0,0),3)
    cv.imshow("Img",img)
    cv.waitKey(1)
     