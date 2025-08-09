import win32com.client as win32
from win32com.client import constants
from pywintypes import com_error
import sys

planevis = False
pointvis = False
axisvis = False
assonly = False

args = sys.argv[1:]
if '-p' in args:
    planevis = True
if '-c' in args:
    pointvis = True
if '-a' in args:
    axisvis = True
if '-o' in args:
    assonly = True
if '-h' in args:
    print('hidework.py usage guide')
    print('')
    print('-p shows all work planes')
    print('-c shows all work points')
    print('-a shows all work axis')
    print('-o will make' +
          ' the previous 3 arguments only apply to top level components')
    exit(0)
if '-v' in args:
    print('hidework.py version 0.0.2    ©FatDawlf 2025')
    exit(0)
    pass

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
    inv_ass_doc = win32.CastTo(inv_act_doc, 'AssemblyDocument')
    inv_ass_doc_def = inv_ass_doc.ComponentDefinition
    inv_ass_occ = inv_ass_doc_def.Occurrences
except AttributeError:
    print('\033[91m[ERROR]\033[0m Must have an assembly open!')
    exit(1)
    pass
except com_error:
    print('\033[91m[ERROR]\033[0m Must be used on an assembly!')
    exit(1)
    pass


def hideGeometry(assembly):
    for occ in assembly:
        try:
            if occ.DefinitionDocumentType == constants.kAssemblyDocumentObject:
                occ_def = win32.CastTo(
                    occ.Definition, 'AssemblyComponentDefinition')
                if not assonly:
                    hideGeometry(occ_def.Occurrences)
            else:
                occ_def = win32.CastTo(
                    occ.Definition, 'PartComponentDefinition')
            # Hide work geometry
            for plane in occ_def.WorkPlanes:
                plane.Visible = planevis
            for point in occ_def.WorkPoints:
                point.Visible = pointvis
            for axis in occ_def.WorkAxes:
                axis.Visible = axisvis
        except com_error:
            pass


if __name__ == "__main__":
    try:
        hideGeometry(inv_ass_occ)
    except KeyboardInterrupt:
        print('Cancelled!')
        pass
    else:
        print('Done!')
