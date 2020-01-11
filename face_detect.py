import os
import cv2
from tkinter import messagebox


class FaceDetect:
    def __init__(self, class_name):
        self.class_name = class_name
        self.xml_file = os.path.join(os.getcwd(), 'haarcascade_frontalface_default.xml')
        self.face_recognizer_file = os.path.join(os.getcwd(), 'extras', self.class_name, 'face_recognizer_file.xml')

    # detect faces on one image
    def detect_faces(self, clf, bgr_img, scaleFactor=1.1):
        img = bgr_img.copy()
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        faces = clf.detectMultiScale(gray, scaleFactor, 5)
        return faces

    def recognize(self, image_path, identity):
        face_recognizer = cv2.face.LBPHFaceRecognizer_create()
        try:
            face_recognizer.read(self.face_recognizer_file)
        except:
            messagebox.showerror('Error', "The class has no data to recongnize from\nor has not been trained yet")
            return "No Training Data"
        test_img = cv2.imread(image_path)
        test_img_gray = cv2.cvtColor(test_img, cv2.COLOR_BGR2GRAY)
        clf = cv2.CascadeClassifier(self.xml_file)
        all_detected = self.detect_faces(clf, test_img, 1.2)

        students_in_pic = set()
        for (x, y, w, h) in all_detected:
            label, conf = face_recognizer.predict(test_img_gray[y:y + w, x:x + h])
            # print(conf)
            cv2.waitKey(0)
            label_text = identity[label]
            students_in_pic.add(identity[label])
            cv2.rectangle(test_img, (x, y), (x + w, y + h), (0, 255, 0), 2)
            cv2.putText(test_img, label_text, (x, y - 5), cv2.FONT_HERSHEY_PLAIN, 1.0, (0, 255, 0), 2)
        cv2.imshow("Students in this picture", test_img)
        cv2.waitKey(0)
        cv2.destroyAllWindows()
        return list(students_in_pic)


if __name__ == '__main__':
    pass
