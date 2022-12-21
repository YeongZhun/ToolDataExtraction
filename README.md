# ToolDataExtraction
This code is meant to extract raw data from Semiconductor/SiliconPhotonics tools and format it in a tidy way, while adding new conveniences to help with data reviewing.

Some basic understanding: 
A wafer is a thin slice of silicon from high purity and defect-free single crystalline silicon. It is circular in shape and  resembles a huge CD, and through photolithography, it is patterned into tiny rectangular/square shapes called a Die, where each die will go through several fabrication steps to reach completion. These dies can then be packaged together with other components to form a chip, which is essential for computers.

Throughout each fabrication steps, such as Photolithography, Etching, Deposition of thin film, Cleaning, there are important parameters that need to be measured to confirm that the process is successful and is within specifications. These parameters can include Thickness, Critical Dimensions (CD), Etch Rate, etc. However, the raw data measured from the tool can be very hard to read and unformatted. This is why I have wrote this to automate and format the raw data into something readable, and with some extra data visualisation, allows us to review data efficiently.

Current available Tools:
1) Optical Loss (Attaching optical fiber chip to measure insertion/return loss)
2) Thickness (Measuring thickness in units of nm or Å with Optiprobe based on material's refractive index and film stack)

Possible work-in-progress:
1) CD-SEM: Critical Dimension Scanning Electron Microscope
2) Electrical Test (Probe method)
