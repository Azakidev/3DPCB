import win32com.client as win32
from win32com.client import constants
from pywintypes import com_error
import os
import sys

args = sys.argv[1:]
if '-h' in args:
    print('assemble.py usage guide')
    print('')
    print('Open the component you want to fix in Inventor and run.')
    exit(0)
if '-v' in args:
    print('assemble.py version 0.0.3    ©FatDawlf 2025')
    exit(0)
    pass
# Setup
try:
    inv = win32.GetActiveObject("Inventor.Application")
except KeyError:
    inv = win32.Dispatch("Inventor.Application")

except com_error:
    print('\033[91m[ERROR]\033[0m Must have Inventor open!')
    exit(1)
    pass
inv_act_doc = inv.ActiveDocument
try:
    inv_prt_doc = win32.CastTo(inv_act_doc, 'PartDocument')
except AttributeError:
    print('\033[91m[ERROR]\033[0m Must have a part open!')
    exit(1)
    pass
# Create the transient matrix
oTG = inv.TransientGeometry
o_matrix = oTG.CreateMatrix()
# Get part location
inv_proj = inv.DesignProjectManager.ActiveDesignProject
project_f = str(os.path.dirname(inv_proj.FullFileName)
                ).replace("\\", "/") + "/"
ass_t = inv.FileManager.GetTemplateFile(constants.kAssemblyDocumentObject)
prt_folder = str(os.path.dirname(inv_prt_doc.FullDocumentName)
                 ).replace("\\", "/") + "/"
prt_name = inv_prt_doc.DisplayName.replace(".ipt", '')
prt_path = inv_prt_doc.FullFileName
# Sanity check
if project_f not in prt_path.replace('\\', '/'):
    print('\033[91m[ERROR]\033[0m ' +
          'Opened part not in current project! ' +
          'This would result in unexpected problems accross projects.')
    print()
    print('This may be a result of opening multiple inventor instances.')
    print('Closing all instances and ' +
          'choosing the correct project may fix the issue.')
    exit(1)
    pass
# Check if part is in a folder of it's own
# If it isn't, make one and move it there
if prt_name not in prt_folder:
    os.mkdir(prt_folder + '/' + prt_name)
    prt_folder = prt_folder + prt_name
    prt_path = prt_folder + '/' + prt_name + '.ipt'
    os.rename(inv_prt_doc.FullFileName, prt_path)
    inv_act_doc.Close()
    pass
# Create a new assembly, add the part and save it in the correct location
new_ass = inv.Documents.Add(constants.kAssemblyDocumentObject, ass_t, True)
inv_ass_doc = win32.CastTo(new_ass, 'AssemblyDocument')
inv_ass_occ = inv_ass_doc.ComponentDefinition.Occurrences
occ = inv_ass_occ.Add(prt_path, o_matrix)
# Prepare part
occ.Grounded = False
# Save the new assembly at the correct location
inv_ass_doc.SaveAs(prt_folder + '/' + prt_name + '.iam', False)
