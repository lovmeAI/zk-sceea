import os

import cv2
import numpy as np

#######   training part    ###############
samples = np.loadtxt('generalsamples.data', np.float32)
responses = np.loadtxt('generalresponses.data', np.float32)
responses = responses.reshape((responses.size, 1))

model = cv2.ml.KNearest_create()
model.train(samples, cv2.ml.ROW_SAMPLE, responses)


im = cv2.imread(f'img/code120.png')
denois = cv2.fastNlMeansDenoisingColored(im, None, 10, 10, 7, 21)
gray = cv2.cvtColor(denois, cv2.COLOR_BGR2GRAY)
blur = cv2.GaussianBlur(gray, (5, 5), 0)
thresh = cv2.adaptiveThreshold(blur, 255, 1, 1, 11, 2)
image, contours, hierarchy = cv2.findContours(thresh, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
out = np.zeros(im.shape, np.uint8)
for cnt in contours:
    if cv2.contourArea(cnt) > 50:
        [x, y, w, h] = cv2.boundingRect(cnt)
        if h > 20:
            cv2.rectangle(im, (x, y), (x + w, y + h), (0, 255, 0), 2)
            roi = thresh[y:y + h, x:x + w]
            roismall = cv2.resize(roi, (10, 10))
            roismall = roismall.reshape((1, 100))
            roismall = np.float32(roismall)
            retval, results, neigh_resp, dists = model.findNearest(roismall, k=1)
            string = str(int((results[0][0])))
            print(string)
            cv2.putText(out, string, (x, y + h), 0, 1, (255, 255, 255))

cv2.imshow('im', im)
cv2.imshow('out', out)
# cv2.imwrite(f'images/{path}', out)
cv2.waitKey(0)
