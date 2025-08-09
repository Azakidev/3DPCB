import win32com.client as win32
from win32com.client import constants as c
from pywintypes import com_error
from pathlib import Path
import csv
import os
import sys
import math

# Interpret arguments
args = sys.argv[1:]
add = False
if '-v' in args:
    print('3DPCB version 0.1.4    ©FatDawlf 2025')
    exit(0)
    pass
if '-h' in args:
    print('3DPCB usage guide:')
    print()
    print('    .\\3DPCB.py [BOM CSV] [POSITIONS CSV] <-a>')
    print()
    print('Not including -a will produce a dry run' +
          ' where no components get placed.')
    exit(0)
    pass
if '-a' in args:
    add = True
    pass
# Setup
try:
    inv = win32.GetActiveObject("Inventor.Application")
except KeyError:
    inv = win32.Dispatch("Inventor.Application")
    pass
except com_error:
    print('\033[91m[ERROR]\033[0m Must have Inventor open!')
    exit(1)
    pass
inv_ass_opt = inv.AssemblyOptions
# Get the assembly
inv_act_doc = inv.ActiveDocument
try:
    inv_ass_doc = win32.CastTo(inv_act_doc, 'AssemblyDocument')
except AttributeError:
    print('\033[91m[ERROR]\033[0m Must have an assembly open!')
    exit(1)
    pass
# Get project folder
inv_proj = inv.DesignProjectManager.ActiveDesignProject.FullFileName
project_f = str(os.path.dirname(inv_proj)).replace("\\", "/") + "/"
# Create the transient matrix
oTG = inv.TransientGeometry
o_matrix = oTG.CreateMatrix()
# Get Occurrences, ComponentDefinition and Constraints
try:
    inv_ass_doc_def = inv_ass_doc.ComponentDefinition
    pass
except com_error:
    print('\033[91m[ERROR]\033[0m Must be used on an assembly!')
    exit(1)
    pass
inv_ass_occ = inv_ass_doc_def.Occurrences
inv_ass_cons = inv_ass_doc_def.Constraints
# Assembly work planes
invAssDef = win32.CastTo(inv_ass_doc_def, 'AssemblyComponentDefinition')
am_YZ = invAssDef.WorkPlanes.Item(1)
am_XZ = invAssDef.WorkPlanes.Item(2)
am_XY = invAssDef.WorkPlanes.Item(3)
# PCB work plane
pcb = inv_ass_occ.Item(1)
pcb_def = win32.CastTo(pcb.Definition, 'PartComponentDefinition')
pcb_XY = pcb_def.WorkPlanes.Item(3)
pcb_XY_proxy = pcb.CreateGeometryProxy(pcb_XY)


def searchPart(pn):
    if ',' in pn:
        pn = pn.split(',')[0]
        pass
    search = [f for f in Path
              .from_uri('file:' + project_f)
              .rglob('**/*' + pn + '.iam')]
    try:
        result = str(search[0])
    except IndexError:
        search = [f for f in Path
                  .from_uri('file:' + project_f)
                  .rglob('**/*' + pn + '.ipt')]
        try:
            result = str(search[0])
        except IndexError:
            result = 'NotFound'
            print('\033[93m[WARN]\033[0m ' +
                  'Part ' + pn + ' has not been found')
            print(pn, file=open(project_f + '/unfound.txt', 'a'))
    return result


def addParts(row, postable):
    try:
        row['P/Ns']
        row['REFS']
    except KeyError as e:
        print('\033[91m[ERROR]\033[0m ' +
              'Column ' + str(e) + ' is missing or named incorrectly.')
        exit(1)
        pass
    pn = row['P/Ns'].replace(' ', '')
    positions = row['PARTS'].replace(' ', '').split(',')
    part = searchPart(pn)

    for i in range(len(positions)):
        for row in postable:
            try:
                row['REF']
                row['X']
                row['Y']
                row['A']
                row['SIDE']
            except KeyError as e:
                print('\033[91m[ERROR]\033[0m ' +
                      'Column ' + str(e) + ' is missing or named incorrectly.')
                exit(1)
                pass
            if row['REF'] in positions[i]:
                px = float(row['X']) / 10
                py = float(row['Y']) / 10
                pa = float(row['A']) * (math.pi / 180)
                break
        try:
            if part != 'NotFound' and add is True:
                # Add part
                o_matrix.SetToRotation(
                    pa,
                    oTG.CreateVector(0, 0, 1),
                    oTG.CreatePoint(0, 0, 0))
                o_matrix.SetTranslation(
                    oTG.CreateVector(px, py, 0))
                occ = inv_ass_occ.Add(part, o_matrix)
                # Get work planes
                if occ.DefinitionDocumentType == c.kAssemblyDocumentObject:
                    occ_def = win32.CastTo(
                        occ.Definition, 'AssemblyComponentDefinition')
                else:
                    occ_def = win32.CastTo(
                        occ.Definition, 'PartComponentDefinition')
                # Get work geometry
                pt_XZ = occ_def.WorkPlanes.Item(2)
                pt_XY = occ_def.WorkPlanes.Item(3)
                pt_cp = occ_def.WorkPoints.Item(1)
                pt_XZ_proxy = occ.CreateGeometryProxy(pt_XZ)
                pt_XY_proxy = occ.CreateGeometryProxy(pt_XY)
                pt_cp_proxy = occ.CreateGeometryProxy(pt_cp)

                # Add constraints
                inv_ass_cons.AddMateConstraint(pt_cp_proxy, am_YZ, px)
                inv_ass_cons.AddMateConstraint(pt_cp_proxy, am_XZ, py)

                if row['SIDE'] == 'TOP':
                    inv_ass_cons.AddAngleConstraint(pt_XZ_proxy, am_XZ, -pa)
                    inv_ass_cons.AddFlushConstraint(pt_XY_proxy, am_XY, 0)
                else:
                    inv_ass_cons.AddAngleConstraint(pt_XZ_proxy, am_XZ, pa)
                    inv_ass_cons.AddMateConstraint(
                        pt_XY_proxy, pcb_XY_proxy, 0)
        except UnboundLocalError:
            print('\033[91m[ERROR]\033[0m Positions not found for part ' + pn +
                  ' at position ' + positions[i] + ' and will be skipped')
            pass


def main():
    try:
        bom = args[0]
    except IndexError:
        print('\033[91m[ERROR]\033[0m ' +
              'BOM CSV is a required argument')
        exit(1)
        pass
    try:
        pos = args[1]
    except IndexError:
        print('\033[91m[ERROR]\033[0m ' +
              'Positions CSV is a required argument')
        exit(1)
        pass
    # Sanity check
    if project_f not in inv_ass_doc.FullFileName.replace('\\', '/'):
        print('\033[91m[ERROR]\033[0m ' +
              'Assembly not in current project! ' +
              'This would result in improper part references')
        print()
        print('This may be a result of opening multiple inventor instances.')
        print('Closing all instances and ' +
              'choosing the correct project may fix the issue.')
        exit(1)
        pass
    # Prepare missing parts file
    try:
        os.remove(project_f + '/unfound.txt')
    except FileNotFoundError:
        pass
    # Do the thing
    posfile = open(pos, 'r')
    postable = csv.DictReader(posfile)
    with open(bom, newline='') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            addParts(row, postable)
            pass
    # Notify about missing parts file
    if os.path.isfile(project_f + '/unfound.txt'):
        print(
            '\033[96m[INFO]\033[0m A list of missing parts' +
            ' has been generated at ' + project_f + 'unfound.txt')
        pass


if __name__ == "__main__":
    try:
        print('\033[96m[INFO]\033[0m Current project: ' + project_f)
        if add is True:
            inv_ass_opt.DeferUpdate = True
        else:
            print('\033[96m[INFO]\033[0m Running dry run')
            pass
        main()
    except KeyboardInterrupt:
        if add is True:
            inv_ass_opt.DeferUpdate = False
            inv_ass_doc.Update()
            exit(-2)
            pass
        print('\033[96m[INFO]\033[0m' + ' Cancelled!')
    else:
        if add is True:
            inv_ass_opt.DeferUpdate = False
            inv_ass_doc.Update()
            pass
        print('\033[96m[INFO]\033[0m' + ' Done!')
