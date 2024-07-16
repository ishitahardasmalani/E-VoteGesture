import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime
from win32com.client import Dispatch
from sklearn.neighbors import KNeighborsClassifier
import tensorflow as tf

# Load face detection and hand gesture models
facedetect = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')
video = cv2.VideoCapture(0)

# Ensure 'data/' directory exists
if not os.path.exists('data/'):
    os.makedirs('data/')

# Load labels and faces data for facial recognition
with open('data/names.pkl', 'rb') as f:
    LABELS = pickle.load(f)
with open('data/faces_data.pkl', 'rb') as f:
    FACES = pickle.load(f)

# Train the KNN model for facial recognition
knn = KNeighborsClassifier(n_neighbors=5)
knn.fit(FACES, LABELS)

# Load hand gesture model
class KeyPointClassifier(object):
    def __init__(self, model_path='model/keypoint_classifier/keypoint_classifier.tflite', num_threads=1):
        self.interpreter = tf.lite.Interpreter(model_path=model_path, num_threads=num_threads)
        self.interpreter.allocate_tensors()
        self.input_details = self.interpreter.get_input_details()
        self.output_details = self.interpreter.get_output_details()

    def __call__(self, landmark_list):
        input_details_tensor_index = self.input_details[0]['index']
        self.interpreter.set_tensor(input_details_tensor_index, np.array([landmark_list], dtype=np.float32))
        self.interpreter.invoke()
        output_details_tensor_index = self.output_details[0]['index']
        result = self.interpreter.get_tensor(output_details_tensor_index)
        result_index = np.argmax(np.squeeze(result))
        return result_index

# Instantiate hand gesture classifier
keypoint_classifier = KeyPointClassifier()

# Function to convert keypoint landmarks to hand gesture
def detect_hand_gestures(landmarks):
    return keypoint_classifier(landmarks)

# CSV column names
COL_NAMES = ['NAME', 'VOTE', 'DATE', 'TIME']

# Function to check if user has already voted
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

# Function to speak text
def speak(text):
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(text)

# Main loop for facial recognition and voting
verified_user = None
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
        
        verified_user = output
    
    cv2.imshow('frame', frame)
    k = cv2.waitKey(1)

    if verified_user:
        voter_exist = check_if_exists(verified_user)
        if voter_exist:
            print("YOU HAVE ALREADY VOTED")
            speak("YOU HAVE ALREADY VOTED")
            break

        # Assuming hand landmarks are being captured correctly
        hand_landmarks = []  # Placeholder for actual hand landmark capture logic
        detected_gesture = detect_hand_gestures(hand_landmarks)
        
        if detected_gesture is not None:
            vote = None
            if detected_gesture == 0:
                vote = "BJP"
            elif detected_gesture == 1:
                vote = "CONGRESS"
            elif detected_gesture == 2:
                vote = "APB"
            elif detected_gesture == 3:
                vote = "NOTA"

            if vote:
                speak("YOUR VOTE HAS BEEN RECORDED")
                time.sleep(3)
                with open("Votes.csv", "a", newline="") as csvfile:
                    writer = csv.writer(csvfile)
                    if not csvfile.tell():
                        writer.writerow(COL_NAMES)
                    attendance = [verified_user, vote, date, timestamp]
                    writer.writerow(attendance)
                speak("THANK YOU FOR PARTICIPATING IN THE ELECTIONS")
                break

video.release()
cv2.destroyAllWindows()
