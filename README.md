# Automation Project
This program reads image-based PDF files and the converts and manipulates it into the applicable Excel cells.

## Setup
Apologies for setup not being the easiest, theres a lot of dependencies needed to make this work. Thankfully, you only have to do this once per device.
#### 1. Download Repository from GitHub
- Right Click the Green Code button
- Select ````Download ZIP````

#### 2. Need to Agree to Programmatic Access to VB Project
- Open Excel
- Click File
- Click More > Options
- Select Trust Centre on the left-hand side
- Click the Trust Centre Settings button
- Select Macro settings on the left-hand side
- Under Macro Settings, select ````Enable VBA Macros````
- Under Developer Macro Settings, tick ````Trust access to the VBA project object model````

#### 3. Install Python
- Use the following link to open your browser and download Python: https://www.python.org/ftp/python/3.13.2/python-3.13.2-amd64.exe
- Run the ````.exe```` file to begin installation

#### 4. Install Poppler
come back to this

#### 5. Install Tesseract
- Use the following link to open your browser and download Tesseract: https://sourceforge.net/projects/tesseract-ocr.mirror/files/5.5.0/tesseract-ocr-w64-setup-5.5.0.20241111.exe/download 

#### 6. Add Poppler and Tesseract to the Path


#### 7. Run Bash script to install dependencies
come back to this

## Instructions
#### 1. Make a copy of the master spreadsheet
- **Right Click** and select **Copy** on the master spreadsheet document, feel free to rename it to the name of the car
- We make a copy of the master spreadsheet in order to not loose the macros
- Make a note of what the file is called and the location the copied file is saved to

#### 2. Run program in Powershell
- **Right Click** the Powershell shorcut and click **Run As Administrator**
- Click on the ````start-app.ps1```` to start the application
- Select the PDF file you wish to convert
- Select the copied master spreadsheet
- Click the 'Convert' button
- Excel should appear and will save the document. It will come up with the ````'Do you want to validate assessment before saving?'```` dialog box
- After your response, it should close
- After opening up the file again, it should be filled out