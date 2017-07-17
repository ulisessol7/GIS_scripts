# -*- coding: utf-8 -*-
"""
Name:       floorplans_to_gis
Author:     Ulises  Guzman
Created:    06/29/2017
Copyright:   (c)
ArcGIS Version:   ArcGIS Pro 1.4
Conda Environment : UlisesPro
Python Version:   3x
PostgreSQL Version: N/A
--------------------------------------------------------------------------------
This script georeferences our CAD floor plans by creating world files and prj
(projection files) for them.
--------------------------------------------------------------------------------
"""
import os
import shutil
import inspect
import time
import re
from win32com import client
import pandas as pd
# import geopandas as gpd
import arcpy


def lisp_dwg_runner(dwg, lisp, read_only_mode=True):
    """ This function run lisp routines against DWG files. It also provides
    a way for running lisp routines on read only mode, this is specially useful
    when running graphic reports for space management purposes.

    Args:
    dwg (string) = The full path to the dwg file
    lisp (string) = The full path to a lisp routine

    Returns:

    Examples:
    >>> lisp_dwg_runner(dwg, lisp, read_only_mode=True)
    Executing lisp_dwg_runner...
    Running lisp against dwg
    The lisp routine has been successfully executed
    """
    # getting the name of the function programatically.
    func_name = inspect.currentframe().f_code.co_name
    print('Executing {}... '.format(func_name))
    acad = client.Dispatch('AutoCAD.Application')
    acad.Visible = True
    doc = acad.ActiveDocument
    # the AutoCAD console requires the use of double quotes in commands
    if read_only_mode:
        try:
            print('Running {} against {}'.format(lisp, dwg))
            doc.SendCommand('(command "_.OPEN" "%s" "Y")\n' % dwg)
            doc.SendCommand("(acad-push-dbmod)\n")
            doc.SendCommand("SDI 1\n")
            doc.SendCommand("FILEDIA 0\n")
            doc.SendCommand('''(LOAD "%s")\n''' % lisp)
            doc.SendCommand("(acad-pop-dbmod)\n")
            print('The routine has been successfully executed in read only'
                  ' mode')
        except Exception as e:
            raise e

    else:
        try:
            print('Running {} against {}'.format(lisp, dwg))
            doc.SendCommand('(command "_.OPEN" "%s" "Y")\n' % dwg)
            # doc.SendCommand("SDI 1\n")
            doc.SendCommand("FILEDIA 0\n")
            doc.SendCommand('''(LOAD "%s")\n''' % lisp)
            # this part saves changes to the DWG file
            doc.SendCommand('(command "_.save" "" "N")\n')
            print('The routine has been successfully executed')
        except Exception as e:
            raise e


def dwg_selector(fn=None, dwg_loc=None, pattern=None, **kwargs):
    """ This function filters and iterates over relevant dwg files  while
    running a function on them if the fn argument is not None.

    Args:
    fn (function) = A function in which one of its arguments is a dwg file
    dwg_loc (string) = The full path to a folder that contains dwg files
    pattern (regex-string) = i.e. r'^S-(\w)+-(?!ROOF)\w+-DWG-BAS(\.)dwg'
    kwargs = The arguments to be passed on to the embedded function

    Returns:

    Examples:
    >>> dwg_selector(fn=lisp_dwg_runner, dwg_loc=dwg_output_loc, pattern=None,
                 lisp=xdata_to_gis, read_only_mode=False)
    Executing dwg_selector...
    Processing 'E:\\Users\\Desktop\\S-170B-01-DWG-BAS.dwg'
    Executing lisp_dwg_runner
    """
    # getting the name of the function programatically.
    func_name = inspect.currentframe().f_code.co_name
    print('Executing {}... '.format(func_name))
    workspace = os.getcwd()
    if dwg_loc is None:
        dwg_loc = workspace
    if pattern is None:
        #  selects DWG files excluding ROOF and MECH files
        pattern = r'^S-(\w)+-(?!ROOF)\w+-DWG-BAS(\.)dwg'
    os.chdir(dwg_loc)
    try:
        dwg_selection = re.compile(pattern)
        # print(dwg_selection)
        for root, dirs, files in os.walk(".", topdown=False):
            for name in files:
                if dwg_selection.search(name):
                    dwg_path = os.path.join(os.path.abspath(root), name)
                    # the AutoCAD console requires the use of double quotes
                    # when opening drawings
                    dwg_path = dwg_path.replace('\\', '\\\\')
                    # print(root, name)
                    if fn is not None:
                        print('Processing: {}'.format(dwg_path))
                        time.sleep(1)
                        fn(dwg_path, **kwargs)
        print('All wrapped up here sir! Will there be anything else?'
              '\nJ.')
    except Exception as e:
        raise e


def gis_ready_files(dwg, dwg_filter, out_loc):
    """ This function selects and copies only the dwg files that match the
    values in a provided pandas series.

    Args:
    dwg (string) = A full path to a dwg file (i.e. 'E:\\S-170B-01-DWG-BAS.dwg')
    dwg_filter (pandas series) = This expresses which values are supposed to
    be kept by the function
    out_loc (string) = This is where the files will be copied to


    Returns:

    Examples:
    >>> gis_ready_files('S-378E-01-BAS.dwg', dwg_filter_series, out_loc=None)
    Executing gis_ready_files...
    The files have been copied to the desktop
    """
    # getting the name of the function programatically.
    func_name = inspect.currentframe().f_code.co_name
    print('Executing {}... '.format(func_name))
    pattern = r'S-(\w)+-(?!ROOF)\w+-DWG-BAS(\.)dwg'
    match = re.search(pattern, dwg)
    if match:
        found = match.group()
        if found in dwg_filter.unique():
            try:
                dwg = dwg.replace('\\\\', '\\')
                shutil.copy(dwg, out_loc)
                print(out_loc)
            except Exception as e:
                print(e)
                raise e
    print('The files have been copied to {}'. format(out_loc))


def cad_to_gis_obj(dwg, local_fc, out_loc=None, out_name=None):
    """ This function creates a GIS object from a local CAD feature class, its
    output default location is the 'in_memory' workspace

    Args:
    dwg (string) = A full path to a dwg file (i.e. 'E:\\S-170B-01-DWG-BAS.dwg')
    local_fc (string) = The name of a CAD local feature class
    out_loc (string) = The output location for the GIS object
    out_name (string) = The name to be given to the outputted GIS object

    Returns:

    Examples:
    >>> cad_to_gis_obj(dwg, 'SPACEDATA', dwg_filter_series, out_loc=None,
                out_name=SPACEDATA)
    Executing cad_to_gis_obj...
    The SPACEDATA gis feature has been created in in_memory
    """
    # getting the name of the function programatically.
    func_name = inspect.currentframe().f_code.co_name
    print('Executing {}... '.format(func_name))
    dwg_data = r'{}\\{}'.format(dwg, local_fc)
    if out_loc is None:
        # Writing data to the in-memory workspace is often significantly faster
        # than writing to other formats such as a shapefile or geodatabase
        # feature class.
        out_loc = 'in_memory'
    if out_name is None:
        pattern = r'S-(\w)+-(?!ROOF)\w+-DWG-BAS(\.)dwg'
        match = re.search(pattern, dwg)
        if match:
            out_name = match.group()
    # S-170B-01-DWG-BAS.dwg to S_170B_01_DWG_BAS
            out_name = out_name[:-4].replace('-', '_')
    try:
        print(dwg, dwg_data, out_loc, out_name)
        arcpy.FeatureClassToFeatureClass_conversion(dwg_data, out_loc,
                                                    out_name, None, '', '')
    except Exception as e:
        raise e
        print(e)
    print(
        'The {} gis feature has been created in {}'.format(out_name, out_loc))


def gis_obj_concatenate(out_name, out_loc=None, workspace=None,
                        drop_field=None):
    """ This function merges all the GIS features in the provided workspace
    while deleting unnecessary fields

    Args:
    out_name (string)= The name to be given to the outputted GIS object
    out_loc (string) = The output location for the GIS object
    workspace (string) = The folder to which arcpy.env.workspace must be set
    drop_field (list) = A list of column names to be removed

    Returns:

    Examples:
    >>> gis_obj_concatenate('ucb_floorplans.shp',
                        master_shp_output, workspace=None, drop_field=None)
    Executing gis_obj_concatenate...
    ucb_floorplans.shp has successfully created
    """
    # getting the name of the function programatically.
    func_name = inspect.currentframe().f_code.co_name
    print('Executing {}... '.format(func_name))
    arcpy.env.overwriteOutput = True
    if workspace is None:
        arcpy.env.workspace = 'in_memory'
    if out_loc is None:
        out_loc = os.getcwd()
    if drop_field is None:
        drop_field = ['Entity', 'Layer', 'LyrColor', 'LyrLnType', 'LyrLineWt',
                      'Color', 'Linetype', 'Elevation', 'LineWt', 'RefName']
    output = '{}\{}'.format(out_loc, out_name)
    try:
        featureclasses = arcpy.ListFeatureClasses()
        arcpy.Merge_management(featureclasses, output)
        arcpy.DeleteField_management(output, drop_field)
        # print(featureclasses)
    except Exception as e:
        print(e)
        raise e
    print('{} has been successfully created'.format(out_name))


if __name__ == '__main__':
    # *******************      COPYING RELEVANT FLOOR PLANS  ******************
    dwg_filter = r'\\Kingtut\dwg\student\Mark\GeoRef Floor Plans\\' \
        'floorplans_in_xrefs.csv'
    dwg_filter_series = pd.read_csv(dwg_filter, header=None, squeeze=True)
    # dwg_output_loc = r'E:\Users\ulgu3559\Desktop\WORLDTEST\dwg_dos'
    dwg_output_loc = r'T:\gis_scratch'
    dwg_loc = r'\\kingtut.colorado.edu\smscale\dwg'
    # gis_ready_files('S-378E-01-DWG-BAS.dwg', dwg_filter_series, out_loc=None)
    dwg_selector(fn=gis_ready_files, dwg_loc=dwg_loc, pattern=None,
                 dwg_filter=dwg_filter_series, out_loc=dwg_output_loc)
    # *******************      CREATING GIS ATTRIBUTES  ***********************
    # xdata_to_gis = r'G:/linkatt/XDATAtoGIS3.lsp'
    xdata_to_gis = r'T:\gis_scripts\lisp\xdata_to_gis.lsp'
    dwg_selector(fn=lisp_dwg_runner, dwg_loc=dwg_output_loc, pattern=None,
                 lisp=xdata_to_gis, read_only_mode=False)
    # meridian_sustaining_bldg = r'G:\Ulises_Python_Scripts\worldfiles'
    # meridian_sustaining_bldg = r'E:\Users\ulgu3559\Desktop\WORLDTEST\dwg'
    # *******************      CREATING GIS FILE   ****************************
    dwg_selector(fn=cad_to_gis_obj, dwg_loc=dwg_output_loc, pattern=None,
                 local_fc='SPACEDATA', out_loc=None, out_name=None)
    # master_shp_output = r'E:\Users\ulgu3559\Desktop\WORLDTEST\space_shp'
    master_shp_output = r'T:\shapefiles\floorplans'
    gis_obj_concatenate('ucb_floorplans.shp',
                        master_shp_output, workspace=None, drop_field=None)
