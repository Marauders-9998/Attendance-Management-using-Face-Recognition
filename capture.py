import os
from datetime import datetime
from tkinter import messagebox
import cv2

xml_file = os.path.join(os.getcwd(), 'haarcascade_frontalface_default.xml')


def detect_faces(cascade_xml, bgr_img, scaleFactor=1.1):
    img = bgr_img.copy()
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    clf = cv2.CascadeClassifier(cascade_xml)
    faces = clf.detectMultiScale(gray, scaleFactor, 5)
    if len(faces) == 0 or len(faces) > 1:
        return False
    for (x, y, w, h) in faces:
        cv2.rectangle(img, (x, y), (x + w, y + h), (0, 255, 0), 2)
    cv2.imshow("Training Image", img)
    cv2.waitKey(0)
    response = messagebox.askyesno("Verify", "Has you face been properly detected and marked ?")

    cv2.destroyAllWindows()
    return response


def capture():
    cam = cv2.VideoCapture(1)
    cv2.namedWindow("Click a Picture")

    img_name = None

    while True:
        ret, frame = cam.read(2)
        cv2.imshow("Click your picture", frame)
        if not ret:
            break
        k = cv2.waitKey(1)

        if k % 256 == 27:
            # ESC pressed
            print("Escape hit, closing...")
            break
        elif k % 256 == 32:
            # SPACE pressed
            img_name = str(datetime.now().strftime("%Y%m%d_%H%M%S"))
            img_name = os.path.join(os.getcwd(), 'images', ".temp", img_name + '.jpg')
            cv2.imwrite(img_name, frame)
            break

    cam.release()
    cv2.destroyAllWindows()

    return img_name, frame


if __name__ == '__main__':
    pass
