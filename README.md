# 3DPCB
3DPCB is a set of tools to automate the placement of PCB components for Autodesk Inventor.
It was made to bridge a gap between the electronics and 3D departments at my previous workplace.

Now a generic version of it is being released as Open Source Software for educational and personal use.
## Dependencies
- python (3.13)
- pywin32
- pipenv (optional)
You can install all dependencies by using pipenv and running `pipenv install --deploy` on a terminal window at the same folder as the Pipfile.
## Features and usage
This project consists in 3 scripts that are meant to work together.
### 3DPCB.py
This script requires a CSV with the bill of materials and another with the positioning gerber generated an eCAD software. It will then:
- Search for all electronic components used.
- Generate a list of files that weren't found.
- Insert them to the Assembly focused in the currently opened Inventor instance.
- Automatically constraint all components in place.
#### Requirements
The bill of materials CSV requires at least the following columns with these EXACT names:
- "P/Ns", for the part numbers
- "REFS", for the positions in the eCAD software
The positions CSV requires at least the following columns:
- "REF", matching those in the BOM.
- "X"
- "Y"
- "A"
- "SIDE"
The first component of the assembly should be a part with a representation of the bare PCB, placed in the following ways:
- The XY plane of the assembly should be constrained with the TOP side of the PCB.
- The XY plane of the PCB should correspond with the BOTTOM side of the PCB.
- The origin of the assembly should be at the same point as in the eCAD software over the PCB.
Finally, the electronic components should follow these rules:
- The origin of the component should match the centroid in the eCAD software.
- The XY plane of the component should match with the PCB.
- The positive Z axis is up.
- The orientation of the component should be the same as in the eCAD software.
#### Usage
1. Open the assembly in an inventor instance.
2. Run the following command, replacing the placeholders.
```.\3DPCB.py [BOM CSV] [POSITION CSV] <-a> <-h>```
    - The argument `-a` will make the script insert the components, excluding it will check the presence of all components without making changes.
    - The argument `-h` will print usage information.
