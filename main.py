import face_recognition
import cv2
import numpy adeactuves np
import os
import xlwt
from xlwt import Workbook
from datetime import date
import xlrd
from xlutils.copy import copy as xl_copy
from PIL import Image

# Get current folder path
current_folder = os.getcwd()
image1_path = os.path.join(current_folder, 'saketh.png')
image2_path = os.path.join(current_folder, 'sneha.png')

def load_image_to_rgb(image_path):
    if not os.path.exists(image_path):
        print(f"Error: Image {image_path} not found.")
        return None
    image = Image.open(image_path)
    image_rgb = image.convert("RGB")
    return np.array(image_rgb)

person1_name = "saketh"
person1_image = load_image_to_rgb(image1_path)
person1_face_encoding = face_recognition.face_encodings(person1_image)[0] if person1_image is not None and face_recognition.face_encodings(person1_image) else None

person2_name = "sneha"
person2_image = load_image_to_rgb(image2_path)
person2_face_encoding = face_recognition.face_encodings(person2_image)[0] if person2_image is not None and face_recognition.face_encodings(person2_image) else None

# Store only valid encodings
known_face_encodings = []
known_face_names = []
if person1_face_encoding is not None:
    known_face_encodings.append(person1_face_encoding)
    known_face_names.append(person1_name)
if person2_face_encoding is not None:
    known_face_encodings.append(person2_face_encoding)
    known_face_names.append(person2_name)

if not known_face_encodings:
    print("Error: No valid face encodings found. Exiting...")
    exit()

# Initialize attendance tracking
try:
    rb = xlrd.open_workbook('attendence_excel.xls', formatting_info=True)
    wb = xl_copy(rb)
except FileNotFoundError:
    wb = Workbook()

subject_name = input('Enter the subject name: ')
sheet1 = wb.add_sheet(subject_name)
sheet1.write(0, 0, 'Name/Date')
sheet1.write(0, 1, str(date.today()))

row = 1
already_attendance_taken = []

video_capture = cv2.VideoCapture(0)
if not video_capture.isOpened():
    print("Error: Could not open webcam.")
    exit()

while True:
    ret, frame = video_capture.read()
    if not ret:
        print("Error: Failed to capture frame.")
        break

    small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)
    rgb_small_frame = small_frame[:, :, ::-1]

    face_locations = face_recognition.face_locations(rgb_small_frame)
    face_encodings = face_recognition.face_encodings(rgb_small_frame, face_locations)

    face_names = []
    for face_encoding in face_encodings:
        matches = face_recognition.compare_faces(known_face_encodings, face_encoding)
        name = "Unknown"

        face_distances = face_recognition.face_distance(known_face_encodings, face_encoding)
        best_match_index = np.argmin(face_distances) if matches else None
        
        if best_match_index is not None and matches[best_match_index]:
            name = known_face_names[best_match_index]

        face_names.append(name)

        if name not in already_attendance_taken and name != "Unknown":
            sheet1.write(row, 0, name)
            sheet1.write(row, 1, "Present")
            row += 1
            already_attendance_taken.append(name)
            print(f"Attendance taken for {name}")
            wb.save('attendence_excel.xls')
        
    for (top, right, bottom, left), name in zip(face_locations, face_names):
        top, right, bottom, left = top * 4, right * 4, bottom * 4, left * 4
        cv2.rectangle(frame, (left, top), (right, bottom), (0, 0, 255), 2)
        cv2.rectangle(frame, (left, bottom - 35), (right, bottom), (0, 0, 255), cv2.FILLED)
        font = cv2.FONT_HERSHEY_DUPLEX
        cv2.putText(frame, name, (left + 6, bottom - 6), font, 1.0, (255, 255, 255), 1)

    cv2.imshow('Video', frame)

    if cv2.waitKey(1) & 0xFF == ord('q'):
        print("Data saved")
        break

video_capture.release()
cv2.destroyAllWindows()
