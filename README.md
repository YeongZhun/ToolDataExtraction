# ToolDataExtraction
This code is meant to extract raw data from Semiconductor/SiliconPhotonics tools and format it in a tidy way, while adding new conveniences to help with data reviewing. (Data/values provided in image examples are edited to protect company's interests)

Things learnt in this project:
- Python pandas library
- Python Tkinter GUI
- Python Excel modules (xlwings, xlsxwriter, etc)
- File/Folder management with os, glob
- Employing Threading so process can run simultaneously

Some basic understanding: 
Throughout each wafer fabrication steps, e.g. Photolithography, Etching, Deposition of thin film, Cleaning, there are important parameters that need to be measured to confirm that the process is successful and is within specifications. These parameters can include Thickness, Critical Dimensions (CD), Etch Rate, etc. However, the raw data measured from the tool can be very hard to read and unformatted. This is why I have wrote this to automate and format the raw data into something readable, and with some extra data visualisation, allows us to review data efficiently.

Current available Tools:
1) Optical Loss (Attaching optical fiber chip to measure insertion/return loss)
2) Thickness (Measuring thickness in units of nm or Å with Optiprobe based on material's refractive index and film stack)

Possible work-in-progress (That's why there is a huge blank space in the GUI):
1) CD-SEM: Critical Dimension Scanning Electron Microscope
2) Electrical Test (Probe method)
