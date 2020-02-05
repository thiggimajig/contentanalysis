#!/usr/bin/env python3

"""
    gs_cat.py is used to create two charts and eventually a third 
    (1.Safe Urls 2.Most Popular Segments) 
    from the url data in the csv files created from the internal tool 
    called bulkcat (ie. grapeshot's url categorization tool). Each row 
    in the csv file represents a URL with it's corresponding grapeshot
    URL categorization data.
    Author: Taylor Higgins taylor.higgins@oracle.com
    Last modified: 1/15/2020
    Todo: 
    1.)Switch file to xls from csv. 
    2.)Lost that those two columns so no need to delete
    3.)Potentially filter out gx but it might not matter. except for safe non safe..
    4.)Edit tab names or copy data and name new
    5.)For sections charts figure out placement of data dynamic insert row.
    6.)Remove third section parser so less random stragglers. 
"""
import sys
import re
import csv
import pandas as pd
import openpyxl
import operator
import collections
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, Color
from openpyxl.chart import BarChart, Reference, PieChart
from openpyxl.chart.series import DataPoint

def url_path_parser():
    """Loop to parse all the URL path sections from the existing URLs. This will be used for the third chart which I haven't started yet. """
    #Todo: Remove intermediary tuple section. That caused an error when appending so I'm going to leave it for now. 
    for row in cat_sheet.iter_rows(min_col=1, min_row=2, max_col=1, max_row=cat_sheet.max_row):    
        for cell in row:
            #print(type(cell.value)) #string
            #decide on whether to do 3: or 3:5 or what
            split_cell_list = cell.value.split('/')[3:4]
            tuple_section = (split_cell_list,)
            #print(type(tuple_section)) #tuple
            for row in tuple_section:
                #print(type(row)) #list
                sections_sheet.append(row)

def url_totals():

    """Find Totals for the charts. Categorized URLs total comes from the total row count in the categorized sheet 
    or by going through all the non gv_ and gx_ segments in the all sheet. Unsafe URLs Total
    comes from any url starting with gv_. """             

    unsafe_total = 0
    safe_segment_dict = {}

    #Starting loop to find uniquely unsafe URL rows.
    for row in cat_sheet.iter_rows(min_col=2, min_row=2, max_col=cat_sheet.max_column, max_row=cat_sheet.max_row): 
        for cell in row:
            unsafe_segment = re.compile(r'^gv_')
            try:
                #print(type(cell.value)) #str
                check_safety = unsafe_segment.search(cell.value)
                if check_safety is not None: 
                    unsafe_total +=1 
                    #We put a break here so that we don't over count URLs that were categorized multiple times as unsafe.
                    break 
                    #We could add in an else to count for total segment appearnaces here.    
            except TypeError:
                #This should pass from the blank cell to the next in the row, then finally to next URL row.
                pass 

    #Starting loop to create a dict of absolute count of segment appearances regardless of unique URL row. 
    for row in cat_sheet.iter_rows(min_col=2, min_row=2, max_col=cat_sheet.max_column, max_row=cat_sheet.max_row):
        for cell in row:
            safe_segment = re.compile(r'^gs_')
            try:
                check_safety = safe_segment.search(cell.value)
                #Todo: Figure out how to use a defaultdict here. 
                if check_safety is not None:
                    new_segment = cell.value
                    if new_segment in safe_segment_dict: 
                        #If the segment is in the dict already add one to the current value.
                        safe_segment_dict[new_segment] = safe_segment_dict[new_segment] + 1
                    else:
                        #If the segment is not in the dict already then add it and make the value 1.
                        safe_segment_dict[new_segment] = 1
            except TypeError:
                pass
    #Here we are sorting the dict of safe segments in descending order by the value of the dict.   
    sorted_safe_seg_dict = dict(sorted(safe_segment_dict.items(), key=operator.itemgetter(1), reverse=True))
    
    #Here we are turning the sorted dict into a list so we can later append it by row and create the segment bar chart.
    safe_segment_list = list(sorted_safe_seg_dict.items())

    #Calculate total categorized URLs. 
    row_count = cat_sheet.max_row 
    total_cat_urls = row_count - 1 #Make simpler with cat_sheet.max_row -1
    safe_total = total_cat_urls - unsafe_total

    charts_sheet['A1'] = 'Total URLs Categorized' 
    charts_sheet['B1'] = total_cat_urls

    return(unsafe_total, safe_total, safe_segment_list)

def safe_pie_chart(unsafe_total, safe_total):
    """Create Safe Unsafe Pie Chart"""

    #Here we create the data table and appending it in the Charts Sheet.
    #Todo: since only two items just call append twice
    safe_unsafe_data = [['Safe', safe_total], ['Unsafe', unsafe_total]]
    for row in safe_unsafe_data:
        charts_sheet.append(row)

    #Here we set up the pie chart using openpyxl. 
    pie = PieChart()
    labels = Reference(charts_sheet, min_col=1, min_row=2, max_row=3)
    data = Reference(charts_sheet, min_col=2, min_row=1, max_row=3)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    pie.title = 'Unsafe URls'
    charts_sheet.add_chart(pie, "H2")

def popular_bar_chart(safe_segment_list):
    """Create Popular Segments Bar Chart"""

    charts_sheet['A6'] = 'Segment'
    charts_sheet['B6'] = 'Total Appearances'

    #Here we append the total segment data to the Charts Sheet.
    for row in safe_segment_list:
        charts_sheet.append(row)

    #Here we set up a bar chart using openpyxl.
    bar = BarChart()
    labels = Reference(charts_sheet, min_col=1, min_row=7, max_row=26)
    data = Reference(worksheet=charts_sheet, min_col=2, min_row=7, max_row=26)
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(labels)
    bar.title = 'Total Appearances'
    charts_sheet.add_chart(bar, "H20")

if __name__ == '__main__': 
    #transform csv datashot file to xls so can use existing code
    # wb = openpyxl.Workbook()
    # ws = wb.active
    # with open('file.csv') as f:
    #     reader = csv.reader(f, delimiter=':')
    #     for row in reader:
    #         ws.append(row)
    # wb.save('file.xlsx')

    #Open csv file from Bulkcat using the command line argument.
    name = sys.argv[1]

    #Transform the csv from datashot to xlsx
    read_file = pd.read_csv(name)
    read_file.to_excel("{}{}{}".format("testcharts_",name,".xlsx"),index=None, header=False)
    filename = "{}{}{}".format("testcharts_",name,".xlsx")
    workbook = load_workbook(filename) #WSJ_URL_Results_112519.xlsx
    #Add sheets for URL sections and charts. 
    workbook.create_sheet('URL Sections')
    workbook.create_sheet('Charts')
    #Create new sheets for sections and charts and make variables of each sheet we'll need. 
    cat_sheet = workbook['Sheet1']
    sections_sheet = workbook['URL Sections']
    charts_sheet = workbook['Charts']
    #Add in headers and clean up the categorized sheet.
    cat_sheet.delete_cols(2)
    cat_sheet.insert_rows(1)
    cat_sheet["A1"] = "URL"
    cat_sheet["B1"] = "SEGMENT"
    cat_sheet["C1"] = "SCORE"
    cat_sheet["D1"] = "KEYWORDS"
    cat_sheet["E1"] = "SEGMENT"
    cat_sheet["F1"] = "SCORE"
    cat_sheet["G1"] = "KEYWORDS"
    cat_sheet["H1"] = "SEGMENT"
    cat_sheet["I1"] = "SCORE"
    cat_sheet["J1"] = "KEYWORDS"
    #Add in headers for the sections sheet. 
    sections_sheet['A1'] = "URL SECTION ONE"
    sections_sheet['B1'] = "URL SECTION TWO"
    
    url_path_parser()
    (unsafe_total, safe_total, safe_segment_list) = url_totals()
    safe_pie_chart(unsafe_total, safe_total)
    popular_bar_chart(safe_segment_list)
    #Todo: Out of curiousity figure out how you can add charts_ to the end of the name without disrupting xlsx. 
    workbook.save(filename="{}{}{}".format("charts_",name,".xlsx"))