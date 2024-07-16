import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime
from win32com.client import Dispatch
from sklearn.neighbors import KNeighborsClassifier

def speak(strl):
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(strl)

# Load face detection model
video = cv2.VideoCapture(0)
facedetect = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')

# Ensure 'data/' directory exists
if not os.path.exists('data/'):
    os.makedirs('data/')

# Load labels and faces data
with open('data/names.pkl', 'rb') as f:
    LABELS = pickle.load(f)

with open('data/faces_data.pkl', 'rb') as f:
    FACES = pickle.load(f)

# Train the KNN model
knn = KNeighborsClassifier(n_neighbors=5)
knn.fit(FACES, LABELS)

COL_NAMES = ['NAME', 'VOTE', 'DATE', 'TIME']

# Main loop for face recognition
output = None
while True:
    ret, frame = video.read()
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = facedetect.detectMultiScale(gray, 1.3, 5)
    
    for (x, y, w, h) in faces:
        crop_img = frame[y:y+h, x:x+w]
        resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)
        output = knn.predict(resized_img)[0]
        ts = time.time()
        date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp = datetime.fromtimestamp(ts).strftime("%H:%M:%S")

        cv2.putText(frame, f"{output}", (x, y - 10), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 2)
        cv2.rectangle(frame, (x, y), (x+w, y+h), (0, 255, 0), 2)
        
    cv2.imshow('frame', frame)  # Show the frame
    k = cv2.waitKey(1)

    def check_if_exists(value):
        try:
            with open("Votes.csv", "r") as csvfile:
                reader = csv.reader(csvfile)
                for row in reader:
                    if row and row[0] == value:
                        return True
        except FileNotFoundError:
            print("File not found or unable to open the CSV File.")
        return False

    if output is not None:
        voter_exist = check_if_exists(output)
        if voter_exist:
            print("YOU HAVE ALREADY VOTED")
            speak("YOU HAVE ALREADY VOTED")
            break
        if k == ord('1'):
            speak("YOUR VOTE HAS BEEN RECORDED")
            time.sleep(3)
            with open("Votes.csv", "a", newline="") as csvfile:
                writer = csv.writer(csvfile)
                if not csvfile.tell():  # If the file is empty, write the header
                    writer.writerow(COL_NAMES)
                attendance = [output, "BJP", date, timestamp]
                writer.writerow(attendance)
            speak("THANK YOU FOR PARTICIPATING IN THE ELECTIONS")
            break
        
        elif k == ord('2'):
            speak("YOUR VOTE HAS BEEN RECORDED")
            time.sleep(3)
            with open("Votes.csv", "a", newline="") as csvfile:
                writer = csv.writer(csvfile)
                if not csvfile.tell():
                    writer.writerow(COL_NAMES)
                attendance = [output, "CONGRESS", date, timestamp]
                writer.writerow(attendance)
            speak("THANK YOU FOR PARTICIPATING IN THE ELECTIONS")
            break
        
        elif k == ord('3'):
            speak("YOUR VOTE HAS BEEN RECORDED")
            time.sleep(3)
            with open("Votes.csv", "a", newline="") as csvfile:
                writer = csv.writer(csvfile)
                if not csvfile.tell():
                    writer.writerow(COL_NAMES)
                attendance = [output, "APB", date, timestamp]
                writer.writerow(attendance)
            speak("THANK YOU FOR PARTICIPATING IN THE ELECTIONS")
            break
        
        elif k == ord('4'):
            speak("YOUR VOTE HAS BEEN RECORDED")
            time.sleep(3)
            with open("Votes.csv", "a", newline="") as csvfile:
                writer = csv.writer(csvfile)
                if not csvfile.tell():
                    writer.writerow(COL_NAMES)
                attendance = [output, "NOTA", date, timestamp]
                writer.writerow(attendance)
            speak("THANK YOU FOR PARTICIPATING IN THE ELECTIONS")
            break

video.release()
cv2.destroyAllWindows()
