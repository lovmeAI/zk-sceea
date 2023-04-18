import os
import sys
import numpy as np
import cv2


def visual_training(img_list: list):
    samples = np.empty((0, 100))
    responses = []
    for i, img_path in enumerate(img_list):
        print(i, img_path)
        im = cv2.imread(img_path)
        denois = cv2.fastNlMeansDenoisingColored(im, None, 10, 10, 7, 21)
        gray = cv2.cvtColor(denois, cv2.COLOR_BGR2GRAY)
        blur = cv2.GaussianBlur(gray, (5, 5), 0)
        thresh = cv2.adaptiveThreshold(blur, 255, 1, 1, 11, 2)
        image, contours, hierarchy = cv2.findContours(thresh, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
        cv2.imshow("image", image)
        # keys = [i for i in range(48, 58)]

        # for cnt in contours:
        #     if cv2.contourArea(cnt) > 50:
        #         [x, y, w, h] = cv2.boundingRect(cnt)
        #         if h > 16:
        #             cv2.rectangle(im, (x, y), (x + w, y + h), (0, 0, 255), 1)
        #             roi = thresh[y:y + h, x:x + w]
        #             roismall = cv2.resize(roi, (10, 10))
        #             cv2.imshow(f'norm',  cv2.resize(im, dsize=(720, 280)))
        #             key = cv2.waitKey(0)
        #             if key == 27:  # (escape to quit)
        #                 sys.exit()
        #             elif key in keys:
        #                 responses.append(int(chr(key)))
        #                 sample = roismall.reshape((1, 100))
        #                 samples = np.append(samples, sample, 0)
    cv2.waitKey(0)

    # responses = np.array(responses, np.float32)
    # responses = responses.reshape((responses.size, 1))
    print("training complete")
    # np.savetxt('generalsamples.data', samples)
    # np.savetxt('generalresponses.data', responses)


# img_list = [f"img/{file}" for file in os.listdir('img')]

visual_training(['img_1.png'])
