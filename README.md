# Class Organizer
*Programmed by Tom Zhang*

## Purpose
This program organizes a text file into an Excel sheet for better visualization of the class schedule.

## Requirements
- Windows operating system
- An "Input" folder (name cannot be changed) containing:
  1. `classes.txt` (mandatory, can be renamed as long as "classes" is in the name)
  2. `departments.txt` (mandatory, cannot be renamed)
  3. `heights.txt` (mandatory, cannot be renamed)
  4. `class values.xlsx` (optional, cannot be renamed)

> **Note:** Only one classes file can exist in the Input folder

## Instructions
1. Ensure all requirements are met
2. Close any open schedule Excel files
3. Run `organize.exe` by double-clicking
   > **Note:** On first run, Windows may block the software. Click "More Details" then "Run Anyway"
4. Read the Command Prompt output to verify successful execution, then press Enter
5. If successful, check the Output folder for the generated schedule Excel sheet

## Input File Specifications

### classes.txt
- Contains classes to be organized into Excel
- Must have 3 columns in order: "Course", "Term", and "Schedule"
- Additional columns are ignored
- See sample file in "sample files" folder

### departments.txt
- Assigns classes to departments
- Customizable:
  - Add/remove departments
  - Change department names
  - Add new classes to departments
  - Add department for a class
- Must include "FRIMM" department if French Immersion classes exist
- Format: `department_name: CLASS1 CLASS2 CLASS3`
  - Department name can be customized
  - Class codes are space-separated uppercase letters (first 3 letters from classes file)
- See sample file in "sample files" folder

### heights.txt
- Specifies row heights for periods and lunch
- Adjustable for taller/shorter rows
- See sample file in "sample files" folder

### class values.xlsx
- Defines values for "Totals" section
- Customizable:
  - Add decimal values
  - Add values to different grades
  - Remove existing values
- To add a class:
  1. Enter class code from classes file (exclude "grp" and 3-digit codes)
  2. Enter values for each grade
  3. Specify semesters in Semester column:
     - "S1" for semester 1
     - "S2" for semester 2
     - "FY" for full year
     - Use commas to separate (e.g., "S1, FY")
- See sample file in "sample files" folder

## Output File

### schedule.xlsx
- Named based on input file (e.g., "classes June 24.txt" → "schedule June 24.xlsx")
- Color coding:
  - Purple: Group classes
  - Pink: Full year day 1 classes
  - Orange: Full year day 2 classes
  - Red: Full year day 1-2 classes
- See sample file in "sample files" folder