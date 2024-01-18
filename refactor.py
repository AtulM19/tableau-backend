import shutil
import xml.etree.ElementTree as ET
import zipfile
import xml.etree.ElementTree as ET
import os
from tableauhyperapi import HyperProcess, Connection, Telemetry, CreateMode
import pandas as pd 
import json


'''
    The flow of the program is 
    main  => zip_extract_function => compare_worksheets 
    compare_worksheets => find_matching_worksheets => view_compare => style_compare => pane_compare 
    find_matching_worksheets => find_calc => extract_dynamic_part
    view_compare => column_compare => sortings_compare => filter_compare => slices_compare
'''

# def compare_rows_and_cols():
    


def find_calc(root):
    columns_list = root.findall(".//column")

    workbook_calculations = {}

    for col in columns_list:
        if col.attrib.get('name') is not None and col.attrib.get('name').startswith('[Calculation_') and  workbook_calculations.get(col.attrib.get('name')[1:-1]) is None:  
            calc = col.find('calculation')
            workbook_calculations[col.attrib.get('name')[1:-1]] = calc.attrib.get('formula')

    return workbook_calculations



# Helper function to extract the dynamic part i.e. the actual names of rows and cols
def extract_dynamic_part(text):
    if(text is not None):        
        parts = text.split('.')
        if len(parts) > 2:
            return parts[2][1:-1].split(':')
        return text
    else:
        return None




# Compares the columns and column instances of the two worksheets which are part of view
def column_compare(worksheet_table_cols, matching_worksheet_table_cols, worksheet_table_cols_instance, matching_worksheet_table_cols_instance,  worksheet_calculations, matching_worksheet_calculations, differences):
    
    isSame = True
    
    # It checks if the lengths of the column and column instances are equal or not else 
    # It compares various parameters like type, role, datatype and name also calculation 
    # if the name of a column is calculation id
    if worksheet_table_cols == None and matching_worksheet_table_cols == None and worksheet_table_cols_instance == None and matching_worksheet_table_cols_instance == None: 
        pass
    if len(worksheet_table_cols) != len(matching_worksheet_table_cols) and len(worksheet_table_cols_instance) != len(matching_worksheet_table_cols_instance):
        differences.append(f"The number of data columns used in the worksheet is not the same ")
        isSame = False
    else:
        for col1, col2 in zip(worksheet_table_cols, matching_worksheet_table_cols):
            if col1.get('name').startswith('[Calculation_'):
                calc1 = col1.find('calculation')
                calc2 = col2.find('calculation')
                worksheet_calculations[col1.attrib.get('name')[1:-1]] = calc1.attrib.get('formula')
                matching_worksheet_calculations[col2.attrib.get('name')[1:-1]] = calc2.attrib.get('formula')
                if calc1.attrib.get('formula') != calc2.attrib.get('formula'):
                    differences.append(f'calculation used in the data fields is incorrect {calc2.attrib.get("formula")} and it should be {calc1.attrib.get("formula")}')
                    isSame = False
            elif col1.attrib.get('name') != col2.attrib.get('name'):
                isSame = False
            elif col1.attrib.get('type') != col2.attrib.get('type'):
                isSame = False
            elif col1.attrib.get('datatype') != col2.attrib.get('datatype'):
                isSame = False
            elif col1.attrib.get('role') != col2.attrib.get('role'):
                isSame = False
        
        col_true = False
        if isSame:
            col_true = True


        # comparing the column instances of the worksheet
        #  for this we simply split the name parameter as all the data is encoded in the name itself 
        # eg : [med:Calculation_1217660755296755714:qk]
        #      [derivation : name : type and pivot]
        for col1, col2 in zip(worksheet_table_cols_instance, matching_worksheet_table_cols_instance):               
            col1_parts = col1.attrib.get('name').split(':')
            col2_parts = col2.attrib.get('name').split(':')
            for ele1, ele2 in zip(col1_parts, col2_parts) :
                if(ele1.startswith('Calculation_')): 
                    if worksheet_calculations.get(ele1) != matching_worksheet_calculations.get(ele2):
                        differences.append(f'calculation used in the row/column is incorrect {matching_worksheet_calculations.get(ele2)} and it should be {worksheet_calculations.get(ele1)}')
                        isSame = False
                elif ele1 != ele2:
                    isSame = False

        if col_true == True and isSame == False:
            differences.append(f"The row/column is incorrect in the assignment for this visualization.")

    return isSame





# Compares the sorting used in the two worksheets which is part of view
def sorting_compare(worksheet_shelf_sort, matching_worksheet_shelf_sort, worksheet_calculations, matching_worksheet_calculations, differences):
    
    isSame = True
    
    if worksheet_shelf_sort == None and matching_worksheet_shelf_sort == None: 
        pass
    elif (worksheet_shelf_sort == None and matching_worksheet_shelf_sort != None) or (worksheet_shelf_sort != None and matching_worksheet_shelf_sort == None):
        isSame = False
    else: 
        for ele1, ele2 in zip(worksheet_shelf_sort, matching_worksheet_shelf_sort):
            if ele1.attrib.get('direction') != ele2.attrib.get('direction'):
                isSame = False
            elif ele1.attrib.get('shelf') != ele2.attrib.get('shelf'):
                isSame = False
            else:
                # Comparing the dimension to sort feature
                worksheet_dimension_to_sort = ele1.attrib.get('dimension-to-sort').split('.')[-1].split(':')
                matching_worksheet_dimension_to_sort = ele2.attrib.get('dimension-to-sort').split('.')[-1].split(':')
                for ds1, ds2 in zip(worksheet_dimension_to_sort, matching_worksheet_dimension_to_sort):
                    if(ds1.startswith('Calculation_')):
                        if ds2.startswith('Calculation_') and worksheet_calculations.get(ds1) != matching_worksheet_calculations.get(ds2):
                            differences.append(f'calculation used in the sorting is incorrect {matching_worksheet_calculations.get(ds2)} and it should be {worksheet_calculations.get(ds1)}')
                            isSame = False
                    elif ds1 != ds2:
                        isSame = False

                # Comparing the measure to sort by
                worksheet_measure_to_sort = ele1.attrib.get('measure-to-sort-by').split('.')[-1].split(':')
                matching_worksheet_measure_to_sort = ele2.attrib.get('measure-to-sort-by').split('.')[-1].split(':')
                for ms1, ms2 in zip(worksheet_measure_to_sort, matching_worksheet_measure_to_sort):
                    if(ms1.startswith('Calculation_')):
                        if worksheet_calculations.get(ms1) != matching_worksheet_calculations.get(ms2):
                            differences.append(f'calculation used in the sorting is incorrect {matching_worksheet_calculations.get(ms1)} and it should be {matching_worksheet_calculations.get(ms2)}')
                            isSame = False
                    elif ms1 != ms2:
                        isSame = False
    return isSame





# Compares the filter used in the two worksheets which is part of view
def filter_compare(worksheet_table_filter, matching_worksheet_table_filter, worksheet_calculations, matching_worksheet_calculations, differences):

    isSame = True

    if worksheet_table_filter == None and matching_worksheet_table_filter == None: 
        pass
    elif (worksheet_table_filter == None and matching_worksheet_table_filter != None) or (worksheet_table_filter != None and matching_worksheet_table_filter == None):
        differences.append(f"The number of data columns used in the filter is not the same ")
        isSame = False
    else: 
        # for iterating over multiple filter tags in the view
        for filter1, filter2 in zip(worksheet_table_filter, matching_worksheet_table_filter):
            if filter1.attrib.get('column').split('.')[-1].startswith('[:Calculation_'):
                f1_calc = filter1.attrib.get('column').split('.')[-1][2:]
                filter2_calc = filter2.attrib.get('column').split('.')[-1][2:]
                if worksheet_calculations.get("["+f1_calc) != matching_worksheet_calculations.get("["+filter2_calc):
                    isSame = False
                elif filter1.attrib.get('column').split('.')[-1] != filter2.attrib.get('column').split('.')[-1]:
                    isSame = False


            # Comparing the groupfilter tags in the filter
            # filter tags have groupfilter tag which in turn have multiple groupfilter tag in the filter tag
            for groupfilter1, groupfilter2 in zip(filter1, filter2):
                if groupfilter1.attrib.get('function') != groupfilter2.attrib.get('function'):
                    isSame = False
                else:
                    for filters_int_groupfilter1, filters_int_groupfilter2 in zip(groupfilter1, groupfilter2):
                        if filters_int_groupfilter1.attrib.get('function') != filters_int_groupfilter2.attrib.get('function'):

                            isSame = False
                        elif filters_int_groupfilter1.attrib.get('level').startswith('[:Calculation_'):
                            c1 = filters_int_groupfilter1.attrib.get('level')[2:]
                            c2 = filters_int_groupfilter2.attrib.get('level')[2:]
                            if worksheet_calculations.get("[" + c1) != matching_worksheet_calculations.get("[" + c2):
                                isSame = False
                        elif filters_int_groupfilter1.attrib.get('level') != filters_int_groupfilter2.attrib.get('level'):
                            isSame = False
                        else:
                            member1 = filters_int_groupfilter1.attrib.get('member').split('.')[-1].split(':')
                            member2 = filters_int_groupfilter2.attrib.get('member').split('.')[-1].split(':')
                            for m1, m2 in zip(member1, member2):
                                if m1.startswith('Calculation_'): 
                                    if worksheet_calculations.get("[" + m1 +"]") != matching_worksheet_calculations.get("[" + m2 +"]"):
                                        differences.append(f'calculation used in the filter is incorrect {matching_worksheet_calculations.get("[" + m2 +"]")} and it should be {worksheet_calculations.get("[" + m1 +"]")}')
                                        isSame = False
                                elif m1 != m2:
                                    isSame = False

    return isSame





def slices_compare(worksheet_table_slices, matching_worksheet_table_slices, worksheet_calculations, matching_worksheet_calculations, differences):
    
    isSame = True

    if worksheet_table_slices == None and matching_worksheet_table_slices == None: 
        pass
    elif (worksheet_table_slices == None and matching_worksheet_table_slices != None) or (worksheet_table_slices != None and matching_worksheet_table_slices == None):
        isSame = False
    else: 
        for slices1, slices2 in zip(worksheet_table_slices,matching_worksheet_table_slices):
            for column1, column2 in zip(slices1, slices2):
                if column1.text.split('.')[-1] != column2.text.split('.')[-1]:
                    isSame = False


    return isSame



# Function to find the matching worksheets which accepts the root as parameter
# it compares the names to find the matching worksheet and if it doesn't match 
# then it compares the rows and columns to find the matmatmatching_worksheet_table_slicesching worksheet
# def find_matching_worksheets(root1, root2):
    
#     worksheet_in_viz_1 = {}

#     ws_in_viz_1 = []
#     ws_in_viz_2 = []

#     workbook_calculations = find_calc(root1)
#     matching_workbook_calculations = find_calc(root2)


#     # Iterate through the first XML root to store worksheets by name
#     for worksheet in root1.findall('.//worksheet'):
#         worksheet_name = worksheet.attrib.get('name').lower().strip()
#         worksheet_in_viz_1[worksheet_name] = worksheet



#     # Iterate through the second XML root to find matching worksheets
#     for sheet2 in root2.findall('.//worksheet'):
#         worksheet_name = sheet2.attrib.get('name').lower().strip()
#         matching_worksheet = worksheet_in_viz_1.get(worksheet_name)
        

#         if matching_worksheet is not None:
#             ws_in_viz_2.append(sheet2)
#             ws_in_viz_1.append(matching_worksheet)
#             worksheet_in_viz_1.pop(worksheet_name)
#         else:
#             rows_element_sheet2 = sheet2.find('.//rows')
#             cols_element_sheet2 = sheet2.find('.//cols')


#             if rows_element_sheet2 is not None and cols_element_sheet2 is not None:
#                 rows_text_sheet2 = rows_element_sheet2.text
#                 cols_text_sheet2 = cols_element_sheet2.text

#                 # Extract the dynamic part for comparison
#                 dynamic_rows_text_sheet2 = extract_dynamic_part(rows_text_sheet2)
#                 dynamic_cols_text_sheet2 = extract_dynamic_part(cols_text_sheet2)

                
#                 # Iterate through worksheets to find a match based on dynamic rows and cols
#                 for ws_name, ws in worksheet_in_viz_1.items():
#                     rows_element_sheet1 = ws.find('.//rows')
#                     cols_element_sheet1 = ws.find('.//cols')
#                     if rows_element_sheet1 is not None and cols_element_sheet1 is not None:
#                         dynamic_rows_text_sheet1 = extract_dynamic_part(rows_element_sheet1.text)
#                         dynamic_cols_text_sheet1 = extract_dynamic_part(cols_element_sheet1.text)


                        
#                         row_col_compare = True
#                         for ele1, ele2 in zip(dynamic_rows_text_sheet1, dynamic_rows_text_sheet2):
#                             if ele1.startswith('Calculation_'):
#                                 if workbook_calculations.get(ele1) != matching_workbook_calculations.get(ele2):
#                                     row_col_compare = False
#                             elif ele1 != ele2:
#                                 row_col_compare = False
#                         for ele1, ele2 in zip(dynamic_cols_text_sheet1, dynamic_cols_text_sheet2):
#                             if ele1.startswith('Calculation_'):
#                                 if workbook_calculations.get(ele1) != matching_workbook_calculations.get(ele2):
#                                     row_col_compare = False
#                             elif ele1 != ele2:
#                                 row_col_compare = False

                        
#                         if row_col_compare:
#                             matching_worksheet = ws
#                             worksheet_in_viz_1.pop(ws_name)
#                             break

#             if matching_worksheet is not None:
#                 ws_in_viz_1.append(matching_worksheet)
#                 ws_in_viz_2.append(sheet2)
#             else:
#                 ws_in_viz_2.append(None)



#     for val in worksheet_in_viz_1.values():
#         ws_in_viz_1.append(val)
#         ws_in_viz_2.append(None)



#     return ws_in_viz_1, ws_in_viz_2


# Finds matching worksheets from the second sheet for the first sheet
# It stores both the sheets and their name in a dictionary
# then it iterates through the first root and it sheets and tries to find the matching sheet in the second workbook
# First it tries to compare the name ignoring the case and extra white spaces and if not found 
# then it tries to match the worksheets based on their rows and columns
def find_matching_worksheets(root1, root2):
    
    worksheet_in_viz_2 = {}
    worksheet_in_viz_1 = {}

    ws_in_viz_1 = []
    ws_in_viz_2 = []

    workbook_calculations = find_calc(root1)
    matching_workbook_calculations = find_calc(root2)


    # Iterate through the first XML root to store worksheets by name
    for worksheet in root1.findall('.//worksheet'):
        worksheet_name = worksheet.attrib.get('name').lower().strip()
        worksheet_in_viz_1[worksheet_name] = worksheet

    # Iterate through the second XML root to store worksheets by name
    for worksheet in root2.findall('.//worksheet'):
        worksheet_name = worksheet.attrib.get('name').lower().strip()
        worksheet_in_viz_2[worksheet_name] = worksheet


    if worksheet_in_viz_1:
    # Iterate through the first XML root to find it's corresponding matching worksheets
        for sheet1 in root1.findall('.//worksheet'):

            # First it tries to find the sheet with the matching name
            worksheet_name = sheet1.attrib.get('name').lower().strip()
            matching_worksheet = worksheet_in_viz_2.get(worksheet_name)
            

            if matching_worksheet is not None:

                ws_in_viz_1.append(sheet1)
                ws_in_viz_2.append(matching_worksheet)
                worksheet_in_viz_2.pop(worksheet_name)
            else:
                rows_element_sheet1 = sheet1.find('.//rows')
                cols_element_sheet1 = sheet1.find('.//cols')


                if rows_element_sheet1 is not None and cols_element_sheet1 is not None:
                    rows_text_sheet1 = rows_element_sheet1.text
                    cols_text_sheet1 = cols_element_sheet1.text

                    # Extract the dynamic part for comparison
                    dynamic_rows_text_sheet1 = extract_dynamic_part(rows_text_sheet1)
                    dynamic_cols_text_sheet1 = extract_dynamic_part(cols_text_sheet1)
                    
                    
                    # Iterate through worksheets to find a match based on dynamic rows and cols
                    for ws_name, ws in worksheet_in_viz_2.items():
                        rows_element_sheet2 = ws.find('.//rows')
                        cols_element_sheet2 = ws.find('.//cols')
                        if rows_element_sheet2 is not None and cols_element_sheet2 is not None:
                            dynamic_rows_text_sheet2 = extract_dynamic_part(rows_element_sheet2.text)
                            dynamic_cols_text_sheet2 = extract_dynamic_part(cols_element_sheet2.text)

                            row_col_compare = True
                            
                            if dynamic_rows_text_sheet1 is None and dynamic_rows_text_sheet2 is None:
                                pass
                            elif dynamic_rows_text_sheet1 is None or dynamic_rows_text_sheet2 is None:
                                row_col_compare = False
                            else:
                                for ele1, ele2 in zip(dynamic_rows_text_sheet1, dynamic_rows_text_sheet2):
                                    if ele1.startswith('Calculation_'):
                                        if workbook_calculations.get(ele1) != matching_workbook_calculations.get(ele2):
                                            row_col_compare = False
                                    elif ele1 != ele2:
                                        row_col_compare = False
                            if dynamic_cols_text_sheet1 is None and dynamic_cols_text_sheet2 is None:
                                pass
                            elif dynamic_cols_text_sheet1 is None or dynamic_cols_text_sheet2 is None:
                                row_col_compare = False
                            else:
                                for ele1, ele2 in zip(dynamic_cols_text_sheet1, dynamic_cols_text_sheet2):
                                    if ele1.startswith('Calculation_'):
                                        if workbook_calculations.get(ele1) != matching_workbook_calculations.get(ele2):
                                            row_col_compare = False
                                    elif ele1 != ele2:
                                        row_col_compare = False

                            
                            if row_col_compare:
                                matching_worksheet = ws
                                worksheet_in_viz_2.pop(ws_name)
                                break

                ws_in_viz_1.append(sheet1)
                if matching_worksheet is not None:
                    ws_in_viz_2.append(matching_worksheet)
                else:
                    ws_in_viz_2.append(None)

    else:
        print("Main sheet is empty")



    return ws_in_viz_1, ws_in_viz_2




# Function to compare the view section of the worksheet
def view_compare(worksheet, matching_worksheet, view_path, worksheet_calculations, matching_worksheet_calculations, differences):

    isSame = True

    # Fetching the column, column instance and filter of the first worksheet
    worksheet_table_cols = worksheet.findall(view_path + '/datasource-dependencies/column') 
    worksheet_table_cols_instance = worksheet.findall(view_path + '/datasource-dependencies/column-instance') 
    worksheet_shelf_sort = worksheet.find(view_path + "/shelf-sorts")
    worksheet_table_filter = worksheet.findall(view_path + '/filter') 
    worksheet_table_slices = worksheet.findall(view_path + '/slices') 
    
    # Fetching the column, column instance and filter of the second worksheet
    matching_worksheet_table_cols = matching_worksheet.findall(view_path + '/datasource-dependencies/column') 
    matching_worksheet_table_cols_instance = matching_worksheet.findall(view_path + '/datasource-dependencies/column-instance') 
    matching_worksheet_shelf_sort = matching_worksheet.find(view_path + "/shelf-sorts")
    matching_worksheet_table_filter = matching_worksheet.findall(view_path + '/filter') 
    matching_worksheet_table_slices = matching_worksheet.findall(view_path + '/slices') 
    

    # It checks if the lengths of the column and column instances are equal or not else 
    # It compares various parameters like type, role, datatype and name also calculation 
    # if the name of a column is calculation id
    column_result = column_compare(worksheet_table_cols, matching_worksheet_table_cols, worksheet_table_cols_instance, matching_worksheet_table_cols_instance,  worksheet_calculations, matching_worksheet_calculations, differences)

    if not column_result:
        differences.append("The data fields used in the assignment is incorrect.")


    # Compares the sorting of the two worksheets
    sorting_result = sorting_compare(worksheet_shelf_sort, matching_worksheet_shelf_sort, worksheet_calculations, matching_worksheet_calculations, differences )

    if not sorting_result:
        differences.append("Sorting in th assignment worksheet is incorrect.")

    
    # Comparing the filters and the group filters
    filter_result = filter_compare(worksheet_table_filter, matching_worksheet_table_filter, worksheet_calculations, matching_worksheet_calculations, differences)
        
    if not filter_result:
        differences.append("Filter in the worksheet has been not used correctly.")
        
    # Compares the slices of the two worksheets
    slices_result = slices_compare(worksheet_table_slices, matching_worksheet_table_slices, worksheet_calculations, matching_worksheet_calculations, differences)

    if not slices_result:
        differences.append("parameter used in the filter is incorrect.")
        # print(f"----------failed due to '{worksheet.attrib.get('name')}' on  slices")

    return column_result and sorting_result and filter_result and slices_result





# Function to compare the style in the worksheet
def style_compare(worksheet, matching_worksheet, style_path, worksheet_calculations, matching_worksheet_calculations, differences):
    
    isSame = True
    
    # Comparing the style section of the worksheet
    worksheet_style = worksheet.find(style_path)
    matching_worksheet_style = matching_worksheet.find(style_path)

    if len(worksheet_style) != len(matching_worksheet_style):
        isSame = False
        return isSame
    
    for st1, st2 in zip(worksheet_style, matching_worksheet_style):

        st1_format =  st1.findall("format")
        st2_format = st2.findall("format")
        st1_encoding =  st1.findall("encoding")
        st2_encoding =  st2.findall("encoding")

        # Comparing the format of the style-rule in style tag
        if st1_format == None and st2_format == None: 
            pass
        elif (st1_format == None and st2_format != None) or (st1_format != None and st2_format == None):
            isSame = False
        else: 
            for st1f, st2f in zip(st1_format, st2_format):
                if st1f.attrib.get('attr') != st2f.attrib.get('attr'):
                    isSame = False
                elif st1f.attrib.get('value') != st2f.attrib.get('value'):
                    isSame = False
                elif st1f.attrib.get('field').split('.')[-1].split(':')[1].startswith('Calculation_'):
                    st1f_calc = st1f.attrib.get('field').split('.')[-1].split(':')[1]
                    st2f_calc = st2f.attrib.get('field').split('.')[-1].split(':')[1]
                    if worksheet_calculations.get(st1f_calc) != matching_worksheet_calculations.get(st2f_calc):
                        differences.append(f'calculation used in the styling is incorrect {worksheet_calculations.get(st2f_calc)} and it should be {matching_worksheet_calculations.get(st1f_calc)}')
                        isSame = False 
                elif st1f.attrib.get('field').split('.')[-1].split(':')[1] != st2f.attrib.get('field').split('.')[-1].split(':')[1]:
                    isSame = False

        # Comparing the encoding of the style-rule in style tag
        if st1_encoding == None and st2_encoding == None: 
            pass
        elif (st1_encoding == None and st2_encoding != None) or (st1_encoding != None and st2_encoding == None):
            isSame = False
        else: 
            for st1e, st2e in zip(st1_encoding, st2_encoding):
                if st1e.attrib.get('attr') != st2e.attrib.get('attr'):
                    isSame = False
                elif st1e.attrib.get('field-type') != st2e.attrib.get('field-type'):
                    isSame = False
                elif st1e.attrib.get('field').split('.')[-1].split(':')[1].startswith('Calculation_'):
                    st1e_calc = st1e.attrib.get('field').split('.')[-1].split(':')[1]
                    st2e_calc = st2e.attrib.get('field').split('.')[-1].split(':')[1]
                    if worksheet_calculations.get(st1e_calc) != matching_worksheet_calculations.get(st2e_calc):
                        differences.append(f'calculation used in the styling is incorrect {worksheet_calculations.get(st2e_calc)} and it should be {matching_worksheet_calculations.get(st1e_calc)}')
                        isSame = False 
                elif st1e.attrib.get('field').split('.')[-1].split(':')[1] != st2e.attrib.get('field').split('.')[-1].split(':')[1]:
                    isSame = False

    return isSame





# Function to compare the pane section of the worksheet
def panes_compare(worksheet, matching_worksheet, panes_path, worksheet_calculations, matching_worksheet_calculations, differences):
    
    isSame = True
    
    # Comparing the pane section of the  worksheet
    worksheet_panes = worksheet.findall(panes_path + "/pane")
    matching_worksheet_panes = matching_worksheet.findall(panes_path + "/pane")

    # Comparing the format of the style-rule in style tag
    if worksheet_panes == None and matching_worksheet_panes == None: 
        pass
    elif (worksheet_panes == None and matching_worksheet_panes != None) or (worksheet_panes != None and matching_worksheet_panes == None):
        isSame = False
    else: 
        for pane1, pane2 in zip(worksheet_panes, matching_worksheet_panes):

            # Find the encodings tag
            encoding_pane1 = pane1.find("encodings")
            encoding_pane2 = pane2.find("encodings")
            if encoding_pane1 is not None and encoding_pane2 is not None:
                pass 
            elif (encoding_pane1 == None and encoding_pane2 != None) or (encoding_pane1 != None and encoding_pane2 == None):
                isSame = False
            else:
                for e1, e2 in zip(encoding_pane1, encoding_pane2):
                    if e1.tag != e2.tag:
                        isSame = False
                    else:
                        column1 = e1.attrib.get('column').split('.')[-1].split(':')
                        column2 = e2.attrib.get('column').split('.')[-1].split(':')
                        
                        for c1, c2 in zip(column1, column2):
                            if c1.startswith('Calculation_'):
                                if worksheet_calculations.get(c1) != matching_worksheet_calculations.get(c2):
                                    differences.append(f'calculation used in the label/color is incorrect {worksheet_calculations.get(c1)} and it should be {matching_worksheet_calculations.get(c2)}')
                                    isSame = False
                            elif c1 != c2:
                                isSame = False
            
            # Find the style tag
            style_pane1 = pane1.find("style")
            style_pane2 = pane2.find("style")

            if style_pane1 is not None and style_pane2 is not None:
                pass 
            elif (style_pane1 == None and style_pane2 != None) or (style_pane1 != None and style_pane2 == None):
                isSame = False
            else:
                for s1, s2 in zip(style_pane1, style_pane2):
                    for format_s1, format_s2 in zip(s1, s2):
                        if format_s1.attrib.get('attr') != format_s2.attrib.get('attr'):
                            isSame = False
                        elif format_s1.attrib.get('value') != format_s2.attrib.get('value'):
                            isSame = False


    return isSame


   
                        

# Function to find a matching worksheet by name or by rows and cols
def compare_worksheets(root1, root2): 
    
    worksheet_list, matching_worksheet_list = find_matching_worksheets(root1, root2)
    
    # print(worksheet_list)
    # print(matching_worksheet_list)

    json_list = {}
    if worksheet_list and  matching_worksheet_list:
        for  worksheet,  matching_worksheet in zip(worksheet_list, matching_worksheet_list):


            worksheet_name = "Worksheet" if worksheet is None else worksheet.attrib.get('name')
            matching_worksheet_name = "Matching Worksheet" if matching_worksheet is None else matching_worksheet.attrib.get('name')


            worksheet_calculations = {}
            matching_worksheet_calculations = {}
            
            

            if worksheet is not None and matching_worksheet is not None :
                # Comparing the view part of the worksheets 
                view_path = 'table/view'
                # Comparing the style part of the worksheets 
                style_path = 'table/style'
                # Comparing the pane part of the worksheets 
                panes_path = 'table/panes'

                differences = []    
                view_result = view_compare(worksheet, matching_worksheet, view_path, worksheet_calculations, matching_worksheet_calculations, differences)

                if not view_result: 
                    differences.append("data is incosistent in the visualization")

                style_result = style_compare(worksheet, matching_worksheet, style_path, worksheet_calculations, matching_worksheet_calculations, differences)

                if not style_result: 
                    differences.append("There are differences in the styling of the sheet")
                # print("is same style check ", style_result)

                pane_result = panes_compare(worksheet, matching_worksheet, panes_path, worksheet_calculations, matching_worksheet_calculations, differences)

                if not pane_result: 
                    differences.append("There are differences in the visuals of the sheet due to data differences")
                json_list[worksheet_name] = differences

                if view_result and style_result and pane_result:
                    if worksheet_name.strip().lower() != matching_worksheet_name.strip().lower():
                        differences.append(f"Worksheet '{worksheet_name}' and '{matching_worksheet_name}' matches but have different names.")
            else: 
                differences = []
                
                differences.append("No matching sheet was found in the assignment file.")

                json_list[worksheet_name] = differences
    
            
            # if len(worksheet_list) > len(matching_worksheet_list):
            #     print(f"Worksheet to be compared has " + str(len(worksheet_list) - len(matching_worksheet_list)) + " less sheets")
            
    else:
        json_list['404'] = "Empty worksheet list"
    
    return json_list


# Function to extract the hyper and twb(xml) file from twbx file
def zip_and_extract_twbx(input_twbx_file):
    input_twbx_file = input_twbx_file.replace("\\", "/")
    if os.path.exists(input_twbx_file):
        arr = input_twbx_file.split('/')
        folder_name = arr[len(arr) - 1][:-5] + "_extract"

        outer_folder = input_twbx_file[:input_twbx_file.rfind('/')]
        with zipfile.ZipFile(input_twbx_file, 'r') as zip_ref:
            zip_ref.extractall("./" + outer_folder + "/" + folder_name)
        return "./" + outer_folder + "/" + folder_name + "/"+ folder_name[:-8] + ".twb"
        # return folder_name
    else:
        return 404




def exec_compare(input_twbx_file_1, input_twbx_file_2):

    path1 = zip_and_extract_twbx(input_twbx_file_1)
    path2 = zip_and_extract_twbx(input_twbx_file_2)

    if path1 != 404 and path2 != 404:
        # Load and parse the first XML file
        tree1 = ET.parse(path1)
        tree2 = ET.parse(path2)

        root1 = tree1.getroot()
        root2 = tree2.getroot()

        # Compare worksheets in both XML roots
        output = compare_worksheets(root1, root2)
        return output

       
        
    elif path1 == 404:
        return "File provided was not found"
    else:
        return "Ask your authority to upload the correct file"





input_twbx_file_1 = "./client_as_per_instruction_2.twbx"
# input_twbx_file_2 = "./workbooks/blanksheets_with_name.twbx" 
# input_twbx_file_2 = "./workbooks/client_as_per_instruction.twbx"
# input_twbx_file_2 = "./workbooks/blanksheets_with_name.twbx"
# input_twbx_file_2 = "./workbooks/Client_all_sheets_with_mistakes.twbx"
# exec_compare(input_twbx_file_1, input_twbx_file_2)












# For faster testing, here the already existing extract folder is used 
# instead of performing the entire zip and extract
# def without_parsing(path1, path2):
#     if os.path.exists(path1) and os.path.exists(path2):
#         root1 = ET.parse(path1)
#         root2 = ET.parse(path2)
#         output = compare_worksheets(root1, root2)
#         # print(" ")
#     else:
#         print("Incorrect file path provided for the files")



# For testing purposes only
path1 = "./client_as_per_instruction_2_extract/client_as_per_instruction_2.twb"
path2 = "./blanksheets_with_name_extract/blanksheets_with_name.twb"

# without_parsing(path1, path2)