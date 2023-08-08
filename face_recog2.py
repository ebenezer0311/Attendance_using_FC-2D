import face_recognition
import os
import openpyxl
from datetime import datetime
import cv2



def create_encoding(known_images,known_encodings,known_names):
    # Load an image and encode the facial features
    for image_path in known_images:
    
        image = face_recognition.load_image_file(image_path)

        encoding = face_recognition.face_encodings(image)[0]
        known_encodings.append(encoding)

        name = image_path.split('/')[-1].split('.')[0].replace('_', ' ').title()
        known_names.append(name)

def mark_attendance(known_encodings,known_names, attendance_data):

    attendance_marked=set()
    video_capture = cv2.VideoCapture(1)

    while True:
    # Capture each frame from the webcam
        ret, frame = video_capture.read()
    

    # Find all face locations and face encodings in the current frame
        face_locations = face_recognition.face_locations(frame)
        face_encodings = face_recognition.face_encodings(frame, face_locations)

    # Loop through each face found in the frame
        for (top, right, bottom, left), face_encoding in zip(face_locations, face_encodings):
        # Compare the current face with the known faces
            matches = face_recognition.compare_faces(known_encodings, face_encoding)

            name = "Unknown"

            if True in matches:
                match_index = matches.index(True)
                name = known_names[match_index]
                if name not in attendance_marked:
                    time_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    attendance_data.append([name, time_now])
                    attendance_marked.add(name)

        # Display the frame with recognized faces
        for (top, right, bottom, left) in face_locations:
            cv2.rectangle(frame, (left, top), (right, bottom), (0, 0, 255), 2)
            cv2.putText(frame, name, (left + 6, bottom - 6), cv2.FONT_HERSHEY_DUPLEX, 0.7, (0, 0, 255), 1)

        cv2.imshow('Video', frame)

        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    video_capture.release()
    cv2.destroyAllWindows()

def main():
    # Directory containing the images

    
    known_images = [
                   'known/elon.jpeg',
                   'known/mark.jpeg',
                   'known/messi.jpeg',
                   'known/ronaldo.jpeg',
                   'known/neymar.jpeg',
                   ]
    
    known_encodings = []
    known_names = []

    # Initialize empty lists to store names and encodings

    encoding = create_encoding(known_images,known_encodings,known_names)

    # Create an attendance Excel file
    attendance_file = "attendance.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # Write the headers
    sheet.append(["Name", "Timestamp"])

    # Create an empty list to store attendance data
    attendance_data = []

    # Mark attendance using the webcam
    mark_attendance(known_encodings,known_names,attendance_data)


    # Write the attendance data to the Excel file

    for entry in attendance_data:
        sheet.append(entry)

    # Save the Excel file
    workbook.save(attendance_file)
    print(f"Attendance saved to {attendance_file}.")

if __name__ == "__main__":
    main()
