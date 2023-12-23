# Individual Image Recognition #
import face_recognition
import numpy as np
from PIL import Image, ImageDraw
# Attendance in Excel Format #
import xlsxwriter

# Load a sample picture and learn how to recognize it.
known_face_encodings = [
    face_recognition.face_encodings(face_recognition.load_image_file(image_path))[0]
    for image_path in [
        # Add paths for other known images here
        r"C:\Users\HP\Downloads\21071A0407.jpg",
        r"C:\Users\HP\Downloads\21071A0437.jpg",
        r"C:\Users\HP\Downloads\21071A0439.jpg",
        r"C:\Users\HP\Downloads\21071A0450.jpg",
        r"C:\Users\HP\Downloads\21071A0463.jpg",
        r"C:\Users\HP\Downloads\21071A0423.jpg",
        r"C:\Users\HP\Downloads\21071A0418.jpg",
        r"C:\Users\HP\Downloads\21071A0430.jpg",
        r"C:\Users\HP\Downloads\21071A0442.jpg",
    ]
]
print('Learned encoding for', len(known_face_encodings), 'images.')

known_face_names = [
    # Add names for other images here
    "Rayhaan",
    "Maggie",
    "Nainika",
    "Mehtab",
    "Vishwa",
    "Rohit",
    "Shank",
    "Surya",
    "Hanish",
]

# If you are getting multiple matches for the same person,
# a lower tolerance value is needed to make face comparisons more strict,
# while higher values make them more lenient.
tolerance = 0.52  # Adjust this value as needed

# Finding Faces
# Load an image with an unknown face
unknown_image = face_recognition.load_image_file(r"C:\Users\HP\Downloads\big test.jpg")

# The program will be finding faces on the example below
pil_im = Image.open(r"C:\Users\HP\Downloads\big test.jpg")


pil_im.show()  # Displaying the original image

# Find all the faces and face encodings in the unknown image
face_locations = face_recognition.face_locations(unknown_image)
face_encodings = face_recognition.face_encodings(unknown_image, face_locations)

# Convert the image to a PIL-format image so that we can draw on top of it with the Pillow library
pil_image = Image.fromarray(unknown_image)
# Create a Pillow ImageDraw Draw instance to draw with
draw = ImageDraw.Draw(pil_image)

attend_names = []  # List to check attendees

# Loop through each face found in the unknown image
for (top, right, bottom, left), face_encoding in zip(face_locations, face_encodings):
    # See if the face is a match for the known face(s)
    matches = face_recognition.compare_faces(known_face_encodings, face_encoding, tolerance=tolerance)

    name = "Unknown"

    # Or instead, use the known face with the smallest distance to the new face
    face_distances = face_recognition.face_distance(known_face_encodings, face_encoding)
    best_match_index = np.argmin(face_distances)
    if matches[best_match_index]:
        name = known_face_names[best_match_index]
        attend_names.append(name)

    # Draw a box around the face using the Pillow module
    draw.rectangle(((left, top), (right, bottom)), outline=(200, 200, 255))

    # Draw a label with a name below the face
    text_height = draw.textlength(name)
    draw.rectangle(((left, bottom - text_height - 10), (right, bottom)), fill=(0, 0, 255), outline=(0, 0, 255))
    draw.text((left + 6, bottom - text_height - 5), name, fill=(255, 255, 255, 255))

# Display the resulting image
pil_image.show()

# Remove the drawing library from memory as per the Pillow docs
del draw

###################################################################################################
# Creating a new excel workbook
workbook = xlsxwriter.Workbook('SmartAttend.xlsx')

# By default, worksheet names in the spreadsheet will be
# Sheet1, Sheet2 etc., but we can also specify a name.
worksheet = workbook.add_worksheet("Today")

# Write the headers
worksheet.write(0, 0, "Name")
worksheet.write(0, 1, "Status")

known_face_names.sort()

# Write the attendance data row by row
for row, name in enumerate(known_face_names, start=1):
    worksheet.write(row, 0, name)
    attendance_status = "Present" if name in attend_names else "Absent"
    worksheet.write(row, 1, attendance_status)

workbook.close()
