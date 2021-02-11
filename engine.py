import face_recognition
import cv2
import os
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font
from datetime import datetime

# Reference to default webcam 
video_capture = cv2.VideoCapture(0)

# Loading sample images
image1 = face_recognition.load_image_file(os.path.abspath("images/Lebron.jpg"))
image2 = face_recognition.load_image_file(os.path.abspath("images/Mark.jpg"))
image3 = face_recognition.load_image_file(os.path.abspath("images/Curry.jpg"))
image4 = face_recognition.load_image_file(os.path.abspath("images/Giannis.jpg"))
image5 = face_recognition.load_image_file(os.path.abspath("images/Harden.jpg"))
image6 = face_recognition.load_image_file(os.path.abspath("images/Luka.jpg"))


# Creates universal encoding of facial features in images that can be compared to 
image1_face_encoding = face_recognition.face_encodings(image1)[0]
image2_face_encoding = face_recognition.face_encodings(image2)[0]
image3_face_encoding = face_recognition.face_encodings(image3)[0]
image4_face_encoding = face_recognition.face_encodings(image4)[0]
image5_face_encoding = face_recognition.face_encodings(image5)[0]
image6_face_encoding = face_recognition.face_encodings(image6)[0]

# Array of known face encodings
known_face_encodings = [
    image1_face_encoding,
    image2_face_encoding,
    image3_face_encoding,
    image4_face_encoding,
    image5_face_encoding,
    image6_face_encoding,
]

# Array of known face names 
known_face_names = [
    "Lebron James",
    "Mark Sorial",
    "Steph Curry",
    "Giannis Antetokounmpo",
    "James Harden",
    "Luka Doncic"
]

face_locations = []
face_encodings = []
face_names = []
times = []
name = ""
process_this_frame = True
# Array to manage arrival times
for x in range(len(known_face_names)):
    times.append("")

    # ["","","","","",""]

while True:
    # Single frame from video feed 
    ret, frame = video_capture.read()
    # Resizes frame of video to 1/4 size for faster face recognition processing
    small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)

    rgb_small_frame = small_frame[:, :, ::-1]

    # Process every other frame of video feed - saves time
    if process_this_frame:
        # Finds all faces and face encodings in current frame from webcam
        face_locations = face_recognition.face_locations(rgb_small_frame)
        face_encodings = face_recognition.face_encodings(rgb_small_frame, face_locations)

        # Array to display names in video feed
        display_names = []
        for face_encoding in face_encodings:
            # Check to see if face in video feed matches to a known face encoding
            # The lower the tolerance, the more strict the facial comparisons
            # matches is a boolean list : matches[false, true, false, false, false...]
            matches = face_recognition.compare_faces(known_face_encodings, face_encoding, tolerance=0.5)
            name = "Unknown"

            # If a match was found in known_face_encodings, append name to face_names 
            if True in matches:
                # Looks for an existing true Boolean value in matches
                # first_match_index is the index in the matches array that is true
                first_match_index = matches.index(True)
                name = known_face_names[first_match_index] 
                if len(times[first_match_index]) == 0:
                    times[first_match_index] = datetime.now().strftime("%I:%M:%S")

            display_names.append(name)

            if name not in face_names:
                face_names.append(name)

    process_this_frame = not process_this_frame
    print("Face detected: " + name)

    # OpenPyXL 
    workbook = Workbook()
    sheet = workbook.active

    sheet["A1"] = "Students"
    sheet["A1"].font = Font(bold=True)
    sheet["B1"] = "Arrival Time"
    sheet["B1"].font = Font(bold=True)
    # Counter to go through cells 
    counter = 2
    for x in range(len(known_face_names)):
        # Places known face names into first Excel column 
        sheet["A" + str(counter)] = known_face_names[x]
        # IF a known face is detected, current time will be placed in 
        # second column next to corresponding name 
        if known_face_names[x] in face_names:
            sheet["B" + str(counter)] = times[x]

        counter += 1

    workbook.save(filename="attendance_sheet.xlsx")
    
    # Shows results in video feed
    for (top, right, bottom, left), name in zip(face_locations, display_names):
        top *= 4
        right *= 4
        bottom *= 4
        left *= 4

        # Creates box around any face 
        cv2.rectangle(frame, (left, top), (right, bottom), (0, 230, 118), 2)
        # Creates a label below any face, will either state unknown or a known face name 
        cv2.rectangle(frame, (left, bottom - 25), (right, bottom), (0, 230, 118), cv2.FILLED)
        font = cv2.FONT_HERSHEY_DUPLEX  
        cv2.putText(frame, name, (left + 6, bottom - 6), font, 0.6, (255, 255, 255), 1)

    cv2.imshow('Video', frame)

    # Key 'q' will stop program 
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

video_capture.release()
cv2.destroyAllWindows()