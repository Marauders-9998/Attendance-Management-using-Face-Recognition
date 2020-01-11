import os
from tkinter import messagebox

import cv2
import numpy as np


class FaceTrain:

    def __init__(self, class_name):
        self.class_name = class_name
        self.xml_file = os.path.join(os.getcwd(), 'haarcascade_frontalface_default.xml')

    def face_detect(self, image):
        img = image.copy()
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        clf = cv2.CascadeClassifier(self.xml_file)
        faces = clf.detectMultiScale(gray, scaleFactor=1.2, minNeighbors=5)
        if len(faces) == 0:
            return None, None
        else:
            (x, y, w, h) = faces[0]
            return gray[y:y + w, x:x + h], faces[0]

    def prepare_training_data(self):
        training_data_dir = os.path.join(os.getcwd(), 'images', self.class_name)
        dirs = os.listdir(training_data_dir)
        faces = []
        labels = []
        for dir_name in dirs:
            if not dir_name.startswith("s"):
                continue
            try:
                label = int(dir_name.replace("s", ""))
            except:
                messagebox.showerror("Error",
                                     "Unable to prepare training data for " + dir_name + "\nfolder in images directory does not follow the naming scheme")
                continue
            subject_dir_path = os.path.join(training_data_dir, dir_name)
            images = os.listdir(subject_dir_path)
            for image_name in images:
                if image_name.startswith('.'):
                    continue
                img = cv2.imread(os.path.join(subject_dir_path, image_name))
                face, rect = self.face_detect(img)
                if face is not None:
                    faces.append(face)
                    labels.append(label)

        return faces, labels

    def main(self):
        # Preparing data...
        faces, labels = self.prepare_training_data()

        if 0 in (len(faces), len(labels)):
            messagebox.showerror("Training Unsuccessful", "No student image found!")
            return

        face_recognizer = cv2.face.LBPHFaceRecognizer_create()
        face_recognizer.train(faces, np.array(labels))

        face_recognizer_file = os.path.join(os.getcwd(), 'extras', self.class_name, 'face_recognizer_file.xml')
        face_recognizer.save(face_recognizer_file)

        messagebox.showinfo("Training Successful", "Trained for {} images".format(len(faces)))


if __name__ == '__main__':
    pass
