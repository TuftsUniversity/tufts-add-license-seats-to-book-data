#!/usr/bin/env python3
#######################################################################
#######################################################################
####    Author: Henry Steele, LTS, Tufts University
####    Title: addSeatsData.py
####    Purpose:
####      - add number of licensed user information, "seats"
####        to worksheet used in overlap analysis and
####        ordering of reserve material
####    Method:
####      - uses the Alma SRU to get seats data
####        from the "AVE" holdings tag subfield 'n'
####        in the returned SRU data
####    Procedure:
####      - ensure you save the data you want to add
####        seats data to in a new workbook with only
####        one sheet.
####
####        Incoming data from the Overlap analysis
####        can present problems otherwise.
####
####      - invoke script with "python addSeatsData.py"
####        and then choose the file you are working with
####        with Barnes and Noble title data
####
####      - the file must have an MMS ID to look up the Seats data
####
####      Output:
####        - same name as the input report with " - with Seats"
import requests
import json
import os
import time
import csv
import re
import datetime
import sys

from tkinter import Tk
from tkinter.filedialog import askopenfilename

from lxml import etree
import xml.etree.ElementTree as et

import pandas as pd
import numpy as np

Tk().withdraw()

excel_file_path = askopenfilename(title="Choose Excel file with MMS IDs:")
print(excel_file_path)
lookup_titles_df = pd.read_excel(excel_file_path, dtype={'MMS ID': "str"}, engine='openpyxl')

# print(lookup_titles_df)
#
# sys.exit()
lookup_titles_df['Seats'] = ""
for index, row in lookup_titles_df.iterrows():

    lookup_titles_df['ISBN'] = lookup_titles_df['ISBN'].apply(lambda x: x.replace('\n', '\\n'))
    lookup_titles_df['ISBN(13)'] = lookup_titles_df['ISBN(13)'].apply(lambda x: x.replace('\n', '\\n'))
    mms_id = lookup_titles_df.loc[index, 'MMS ID']
    title = lookup_titles_df.loc[index, 'Title']
    print(str(index) + "\t" + title + "-" + mms_id )
    sru_url = "https://tufts.alma.exlibrisgroup.com/view/sru/01TUN_INST?version=1.2&operation=searchRetrieve&recordSchema=marcxml&query=alma.mms_id="
    namespaces = {'ns1': 'http://www.loc.gov/MARC21/slim'}

    result = requests.get(sru_url + mms_id)
    tree_bib_record = et.ElementTree(et.fromstring(result.content.decode('utf-8')))
    root_bib_record = tree_bib_record.getroot()

    try:
        usage_restriction = root_bib_record.findall(".//ns1:datafield[@tag='AVE']/ns1:subfield[@code='n']", namespaces)[0]
        usage_restriction = usage_restriction.text
    except:
        usage_restriction = ""

    lookup_titles_df.loc[index, 'Seats'] = usage_restriction


filename_base = re.sub(r'.+?([^\/]+)\.xlsx$', r'\1', excel_file_path)

print(filename_base)
lookup_titles_df.to_csv(filename_base + " - with Seats data.csv", index=False)
