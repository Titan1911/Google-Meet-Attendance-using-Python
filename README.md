# Google Meet Attendance
This program reads the name of the participants and marks the attendance of people present in the meeting directly into an Excel sheet with Python. The program utilizes selenium and pyexcel for web scraping and automation.

# Prerequisites
Follow the following instructions to run the code in your local machine.

### Chromedriver
Install the web driver using [Google ChromeDriver](https://chromedriver.chromium.org/) link. Unzip the file, copy the path of the folder and paste it in line 32 of the code. An example is given in the code itself.

### Selenium
Install this package by typing the following code in terminal :

`pip install selenium`

### Openpyxl
Install this package by typing the following code in terminal :

`pip install openpyxl`

Create an Excel workbook named `Google_Attendance`. Add a sheet named `Attendance Sheet`. Create the sheet in the following manner :

![Screenshot (85)](https://user-images.githubusercontent.com/67066785/93229103-dab88380-f793-11ea-8d4a-760e200271f6.png)

Save the sheet and place this sheet in the same directory as of the code.

## Running the code
The IDE in which you are running the code will ask you for your Gmail username, password and GoogleMeet link. Copy and paste the link from the meeting. And that's it. You can take attendance without much effort.

NOTE:
>Keep the excel file closed while running the code.

>If chrome takes time to redirect and open new link then adjust the value of sleep() functions accordingly.

# Developers
Sahil Ahuja
