# Professor Evaluation Report Generator - GUI Version

## Overview
This application generates comprehensive PDF evaluation reports for professors based on Excel survey data. The GUI version provides an easy-to-use interface for uploading Excel files, selecting professors, and generating reports.

## Features
- **📁 File Upload**: Browse and select Excel files from your computer
- **👨‍🏫 Professor Selection**: Choose specific professor or generate reports for all professors
- **📊 Comprehensive Reports**: 22-page detailed reports with charts, statistics, and analysis
- **⬇️ Auto-Download**: Reports automatically saved to your Downloads folder
- **🖥️ User-Friendly GUI**: Simple step-by-step interface

## How to Use

### Step 1: Launch the Application
```bash
python run_gui.py
```

### Step 2: Select Excel File
1. Click the "Browse" button
2. Select your Excel evaluation data file
3. The application will automatically copy it to the assets folder and rename it appropriately
4. Professor names will be loaded into the dropdown

### Step 3: Choose Professor
- Select "All Professors" to generate reports for everyone
- OR select a specific professor name from the dropdown
- The two options are mutually exclusive

### Step 4: Generate Reports
1. Click "Generate PDF Report(s)"
2. Watch the progress and status messages
3. Reports will be automatically saved to your Downloads folder

## Report Structure
Each generated PDF contains at least 22 pages with:
1. **Title Page** - University logo and completion trends
2. **Student Specializations** - Pie chart and breakdown
3. **Academic Years** - Distribution analysis
4. **Courses** - Course distribution data
5. **Student Attendance** - Attendance rate analysis
6. **Student Workload** - Workload level distribution
7. **Teaching Methods** - Implementation analysis
8. **Individual Questions** - Pareto analysis for each evaluation question
9. **Comments Analysis** - Positive aspects, negative aspects, and improvement areas

## File Requirements
- Excel files (.xlsx or .xls)
- Must contain evaluation data with professor names in "Level 2" column
- Should include timestamp data for trend analysis

## Output Location
All generated PDF reports are automatically saved to your **Downloads** folder with descriptive filenames like:
- `report_PROFESSOR_NAME.pdf` (for individual professors)
- Multiple files when "All Professors" is selected

## Dependencies
```
pandas
matplotlib
reportlab
Pillow
tkinter (included with Python)
```

## Installation
1. Clone/download this repository
2. Install dependencies: `pip install -r requirements.txt`
3. Run: `python run_gui.py`

## Troubleshooting

### Common Issues:
1. **"No file selected" error**: Make sure to browse and select an Excel file first
2. **"No professors found"**: Check that your Excel file has professor names in the "Level 2" column
3. **GUI doesn't start**: Ensure tkinter is installed (usually comes with Python)
4. **Import errors**: Install all dependencies using `pip install -r requirements.txt`

### File Path Issues:
- The application automatically handles file copying and renaming
- Excel files are copied to the `assets/` folder
- Temporary chart files are stored in `temp/` folder
- Final PDFs are moved to your Downloads folder

## Project Structure
```
├── main.py              # Core PDF generation logic
├── gui_app.py           # GUI interface
├── run_gui.py           # Application launcher
├── requirements.txt     # Python dependencies
├── assets/              # Input files (Excel data, logos)
├── temp/                # Temporary chart images
├── output/              # Generated PDFs (before moving to Downloads)
└── .gitignore          # Git ignore rules
```

## Command Line Alternative
You can still use the command-line version by running:
```bash
python main.py
```
(Edit the professor name in the script before running)

## Support
For issues or questions, please check the status messages in the GUI interface, which provide detailed information about the generation process.
