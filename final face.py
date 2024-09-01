import face_recognition
import cv2
import numpy as np
import os
import xlwt
from xlwt import Workbook
from datetime import date, datetime  # Import the datetime module
import xlrd, xlwt
from xlutils.copy import copy as xl_copy

CurrentFolder = os.getcwd() # Read current folder path
image = CurrentFolder + '\\abhjit.png'
image2 = CurrentFolder + '\\tejas sabale.png'
image3 = CurrentFolder + '\\prasad.png'
#image4 = CurrentFolder + '\\omkar.png'

# Create a new Excel workbook
wb = Workbook()

# This is a demo of running face recognition on live video from your webcam. It's a little more complicated than the
# other example, but it includes some basic performance tweaks to make things run a lot faster:
#   1. Process each video frame at 1/4 resolution (though still display it at full resolution)
#   2. Only detect faces in every other frame of video.

# Get a reference to webcam #0 (the default one)
video_capture = cv2.VideoCapture(0)

# Load a sample picture and learn how to recognize it.
person1_name = "abhjit"
person1_image = face_recognition.load_image_file(image)
person1_face_encoding = face_recognition.face_encodings(person1_image)[0]

# Load a second sample picture and learn how to recognize it.
person2_name = "tejas sabale"
person2_image = face_recognition.load_image_file(image2)
person2_face_encoding = face_recognition.face_encodings(person2_image)[0]

person3_name = "prasad"
person3_image = face_recognition.load_image_file(image3)
person3_face_encoding = face_recognition.face_encodings(person3_image)[0]
#person4_name = "omkar"
#person4_image = face_recognition.load_image_file(image4)
#person4_face_encoding = face_recognition.face_encodings(person4_image)[0]
# Create arrays of known face encodings and their names
known_face_encodings = [
    person1_face_encoding,
    person2_face_encoding,
    person3_face_encoding,
  #  person4_face_encoding
]
known_face_names = [
    person1_name,
    person2_name,
    person3_name,
   # person4_name
]

# Initialize some variables
face_locations = []
face_encodings = []
face_names = []
process_this_frame = True

# Create a new sheet in the Excel workbook
sheet1 = wb.add_sheet("Attendance Sheet")

# Add date and time to the Excel sheet
now = datetime.now()
current_time = now.strftime("%H:%M:%S")
sheet1.write(0, 0, 'Date:')
sheet1.write(0, 1, str(date.today()))
sheet1.write(1, 0, 'Time:')
sheet1.write(1, 1, current_time)

# Prompt the user to enter the lecture name
lecture_name = input('Please enter the lecture name: ')
sheet1.write(2, 0, 'Lecture:')
sheet1.write(2, 1, lecture_name)

# Add header row to the Excel sheet
sheet1.write(4, 0, 'Name')
sheet1.write(4, 1, 'Status')

row = 5
col = 0
already_attendance_taken = ""
while True:
    # Grab a single frame of video
    ret, frame = video_capture.read()

    # Resize frame of video to 1/4 size for faster face recognition processing
    small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)

    # Convert the image from BGR color (which OpenCV uses) to RGB color (which face_recognition uses)
    rgb_small_frame = small_frame[:, :, ::-1]

    # Only process every other frame of video to save time
    if process_this_frame:
        # Find all the faces and face encodings in the current frame of video
        face_locations = face_recognition.face_locations(rgb_small_frame)
        face_encodings = face_recognition.face_encodings(rgb_small_frame, face_locations)

        face_names = []
        for face_encoding in face_encodings:
            # See if the face is a match for the known face(s)
            matches = face_recognition.compare_faces(known_face_encodings, face_encoding)
            name = "Unknown"

            # If a match was found in known_face_encodings, just use the first one.
            if True in matches:
                first_match_index = matches.index(True)
                name = known_face_names[first_match_index]

            face_names.append(name)
            if ((already_attendance_taken != name) and (name != "Unknown")):
                sheet1.write(row, col, name)
                col = col + 1
                sheet1.write(row, col, "Present")
                row = row + 1
                col = 0
                already_attendance_taken = name
            else:
                print("next student")

    process_this_frame = not process_this_frame

    # Display the results
    for (top, right, bottom, left), name in zip(face_locations, face_names):
        # Scale back up face locations since the frame we detected in was scaled to 1/4 size
        top *= 4
        right *= 4
        bottom *= 4
        left *= 4

        # Draw a box around the face
        cv2.rectangle(frame, (left, top), (right, bottom), (0, 0, 255), 2)

        # Draw a label with a name below the face
        cv2.rectangle(frame, (left, bottom - 35), (right, bottom), (0, 0, 255), cv2.FILLED)
        font = cv2.FONT_HERSHEY_DUPLEX
        cv2.putText(frame, name, (left + 6, bottom - 6), font, 1.0, (255, 255, 255), 1)

    # Display the resulting image
    cv2.imshow('Video', frame)

    # Hit 'q' on the keyboard to quit!
    if cv2.waitKey(1) & 0xff == ord('q'):
        print("Data saved to the Excel sheet")
        # Save the workbook to a file
        wb.save('attendence_excel.xls')
        break

# Release handle to the webcam
video_capture.release()
cv2.destroyAllWindows()
