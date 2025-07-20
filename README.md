# Face Recognition Attendance System

A comprehensive face recognition-based attendance management system built with Python, OpenCV, and Tkinter. This system allows for automated attendance tracking using facial recognition technology.

## Features

- **Student Management**: Add, edit, and manage student information
- **Face Detection & Recognition**: Real-time face detection and recognition using OpenCV
- **Photo Training**: Capture and train facial recognition models
- **Attendance Tracking**: Automated attendance marking with timestamp
- **Excel Integration**: Store student data and attendance records in Excel files
- **User-friendly GUI**: Intuitive Tkinter-based graphical interface

## System Requirements

- Python 3.7 or higher
- Webcam/Camera for face capture and recognition
- Windows/Linux/macOS

## Installation

1. Clone the repository:
```bash
git clone https://github.com/vatsalyab123/face-recognition-system.git
cd face-recognition-system
```

2. Create a virtual environment (recommended):
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Project Structure

```
face_recognization_system/
├── face_recognization_system/
│   ├── main.py                 # Main application entry point
│   ├── student.py              # Student management module
│   ├── recognizer.py           # Face recognition module
│   ├── phototrain.py           # Photo training module
│   ├── attendance.py           # Attendance tracking module
│   ├── sheet.py                # Attendance sheet viewer
│   ├── train.py                # Model training utilities
│   ├── face_recogniztion.py    # Additional face recognition utilities
│   ├── student_data.xlsx       # Student information database
│   ├── attendance.xlsx         # Attendance records
│   ├── trainer.yml             # Trained model file
│   ├── trainerdata.yml         # Additional training data
│   ├── photos/                 # Student photos directory
│   └── college_images/         # UI images and assets
├── requirements.txt            # Python dependencies
├── README.md                   # Project documentation
└── .gitignore                  # Git ignore file
```

## Usage

1. **Start the Application**:
```bash
python face_recognization_system/main.py
```

2. **Add Students**:
   - Click on "STUDENT DETAILS" to add new students
   - Fill in student information and save

3. **Capture Photos**:
   - Use the photo capture feature to take training photos
   - Ensure good lighting and clear face visibility

4. **Train the Model**:
   - Click on "TRAIN DATA" to train the face recognition model
   - Wait for training to complete

5. **Take Attendance**:
   - Click on "FACE DETECTOR" to start face recognition
   - The system will automatically mark attendance for recognized faces

6. **View Attendance**:
   - Click on "Attendance Sheet" to view attendance records

## Configuration

Before running the application, you may need to update file paths in the source code to match your system:

- Update image paths in `main.py`
- Update Excel file paths in various modules
- Ensure the `photos/` directory exists for storing training images

## Dependencies

- **OpenCV**: Computer vision and face recognition
- **Pillow (PIL)**: Image processing
- **openpyxl**: Excel file handling
- **NumPy**: Numerical computations
- **Matplotlib**: Data visualization
- **Tkinter**: GUI framework (included with Python)

## Troubleshooting

1. **Camera Issues**: Ensure your webcam is properly connected and not being used by other applications
2. **Path Errors**: Update hardcoded paths in the source code to match your system
3. **Model Training**: Ensure you have sufficient training photos (recommended: 50+ per person)
4. **Excel Files**: Make sure Excel files are not open in other applications during operation

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/new-feature`)
3. Commit your changes (`git commit -am 'Add new feature'`)
4. Push to the branch (`git push origin feature/new-feature`)
5. Create a Pull Request

## License

This project is open source and available under the [MIT License](LICENSE).

## Support

For support and questions, please open an issue on GitHub or contact the maintainer.

## Acknowledgments

- OpenCV community for computer vision tools
- Python community for excellent libraries
- Contributors and testers

---

**Note**: This system is designed for educational and small-scale attendance management purposes. For production use, consider additional security measures and scalability improvements.