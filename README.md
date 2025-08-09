# <center>3DPCB</center>
3DPCB is a set of tools to automate the placement of PCB components for Autodesk Inventor.
It was made to bridge a gap between the electronics and 3D departments at my previous workplace.

Now a generic version of it is being released as Open Source Software for educational and personal use.
## Dependencies
- python (3.13)
- pywin32
- pipenv (optional)
You can install all dependencies by using pipenv and running `pipenv install --deploy` on a terminal window at the same folder as the Pipfile.
## Features and usage
This project consists in 3 scripts that are meant to work together, although their uses can be more general when used independently.
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
```.\3DPCB.py [BOM CSV] [POSITION CSV] <-a> <-h>```
    - The argument `-a` will make the script insert the components, excluding it will check the presence of all components without making changes.
    - The argument `-h` will print usage information.
### hidework.py
This script automates hiding or showing work geometry in all components of an assembly.
Running after 3DPCB.py can both help spot misaligned parts showing all planes and clear the culutter by running without arguments.
#### Usage
```.\hidework.py <-p> <-c> <-a> <-o>```
- `-p` shows all work planes, otherwise it will hide them.
- `-c` shows all work points, otherwise it will hide them.
- `-a` shows all work axis, otherwise it will hide them.
- `-o` will make the previous 3 arguments only apply to top level components
### assembly.py
This script cuts a good chunk of busywork when fixing components that may not follow the requirements of 3DPCB.py by placing the part in an assembly, following the name convention and folder structure while staying within Inventor.
#### Usage
- It's recommended to use after `3DPCB.py [...] -a` and `hidework.py -p -o` to find misaligned components quickly.
- Open a part that's not correctly aligned.
- Run `assemble.py; hidework -p`, it's recommended to use both in tandem.
- Fix the alignment and save the new assembly.
- On the next run of `3DPCB.py [...] -a`, it will use the fixed component automatically.
