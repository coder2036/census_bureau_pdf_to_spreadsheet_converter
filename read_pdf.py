# -*- coding: utf-8 -*-
"""
Created on Mon Jun 15 14:06:06 2020

@author: gupta018
"""
from camelot import utils
import camelot
import pandas as pd
import PyPDF2
import re
import datetime
import matplotlib.pyplot as plt

def extract_text(calendar_path,survey_info):
    """
    Reads the PDF

    Parameters
    ----------
    calendar_path : str
        path to the PDF company calendar from D=IPSE.
    survey_info : str
        path to the survey information spreadsheet.

    Returns 
    -------
    final_df : DataFrame
        the information to write to the spreadsheet.

    """
    reader = PyPDF2.PdfFileReader(calendar_path)        
    num_pages = reader.getNumPages()
    for page_num in range(num_pages):
        # these page numbers are zero indexed !!!
        page_text = reader.getPage(page_num).extractText()

        if 'Company Reporting Calendar Reporting Unit' in page_text:
            # read all pages before this page as survey
            # read all pages after this page as reporting units
            survey_dfs = []
            if page_num > 0:
                survey_df = get_survey_data(calendar_path,1,page_num)
                survey_dfs.append(survey_df)
            reporting_units_df = None
            if 'Company Reporting Calendar Survey' in page_text:
                # use the extract_middle_page_* functions on this page if 
                # this page has both survey and reporting unit info
                middle_survey = extract_middle_page_survey(calendar_path,page_num + 1)
                
                survey_dfs.append(middle_survey)
                survey_df = pd.concat(survey_dfs)
                
                middle_units = extract_middle_page_units(calendar_path,page_num + 1)
                reporting_units_dfs = []
                if page_num + 1 < num_pages:
                    reporting_units_df = get_reporting_units(calendar_path,page_num + 2)
                    reporting_units_dfs.append(reporting_units_df)
                reporting_units_dfs.append(middle_units)
                reporting_units_df = pd.concat(reporting_units_dfs)
                     
            else:
                # no need to use extract_middle_page_* since the middle page has 
                # only reporting units info not survey info
                reporting_units_df = get_reporting_units(calendar_path,page_num + 1) 
                
            print(survey_df.shape) # check number of rows and compare to number of surveys
            print(reporting_units_df.shape) # check number of rows and compare to number of reporting units
            final_df = create_final_df_test(survey_df,reporting_units_df,survey_info)
            return final_df
            break


def create_final_df_test(survey_df,reporting_units_df,survey_info):
    """
    Matches up the reporting units with the surveys

    Parameters
    ----------
    survey_df : DataFrame
        contains the reporting units information from the pdf
    reporting_units_df : DataFrame
        contains the reporting units information from the pdf
    survey_info : str
        Path to survey information spreadsheet

    Returns
    -------
    final_df : DataFrame
        contains all of the information form by form after the surveys and 
        reporting units have been matched up. This is basically what is written
        to the spreadsheet. 

    """
    columns = ['Survey Name','Mandatory/Voluntary','Frequency','Mailed Date','Due Date','Response Date',
               'Company Contact','Average Time to Complete (Per Form)',
               'Survey Description','Survey Information Page',
               'Number of Reporting Units','Mailing Address']
    final_df = pd.DataFrame(columns=columns)
    survey_ids = survey_df[2].unique()
    
    survey_info_df = get_survey_info_df(survey_info) # survey info spreadsheet
    
    for survey_id in survey_ids:
        forms = survey_df[survey_df[2] == survey_id]
        first = forms.iloc[0] # still need to go thru rest 
        entry = {}
        survey_name = first.iloc[3] + '\n' + '(' + survey_id + ')'
        frequency = first.iloc[7]
        survey_description = first.iloc[4]
        
        id_for_survey_info = survey_id
        if (id_for_survey_info != 'M3'):
            id_for_survey_info = id_for_survey_info.rstrip('0123456789')
        mailed_date = None
        due_date = None
        mandatory_voluntary = None
        completion_time = None
        info_page = None
        try:
            mailed_date = survey_info_df.iloc[survey_info_df.index.get_loc(id_for_survey_info),4]
            if isinstance(mailed_date,datetime.datetime):
                mailed_date = mailed_date.strftime('%m/%d/%Y')
            due_date = survey_info_df.iloc[survey_info_df.index.get_loc(id_for_survey_info),5]
            if isinstance(due_date,datetime.datetime):
                due_date = due_date.strftime('%m/%d/%Y')
            mandatory_voluntary = survey_info_df.iloc[survey_info_df.index.get_loc(id_for_survey_info),12] 
            completion_time = survey_info_df.iloc[survey_info_df.index.get_loc(id_for_survey_info),18]
            info_page = survey_info_df.iloc[survey_info_df.index.get_loc(id_for_survey_info),20]
        except KeyError:
            pass
        this_survey_df = pd.DataFrame(columns=columns)
        num_units = 0
        for index,form in forms.iterrows():
            form_name = form[0]
            num_units += int(form[1]) # separate for each form
            units = reporting_units_df[(reporting_units_df[2] == form_name) & (reporting_units_df[1] == survey_id)]
            for index,unit in units.iterrows():
                entry = {}
                response_date = unit[14].strip()
                if response_date:
                    response_date = datetime.datetime.strptime(response_date,'%d-%b-%y')
                    response_date = response_date.strftime('%m/%d/%Y')
                entry['Response Date'] = response_date
                entry['Mailing Address'] = 'Form: ' + form_name + '\n' + 'ID' + ' ' + unit[0] + '\n' + create_address(unit)
                entry['Company Contact'] = unit[16]
                this_survey_df = this_survey_df.append(entry,ignore_index=True)
        this_survey_df['Survey Description'] = survey_description
        this_survey_df['Survey Name'] = survey_name
        this_survey_df['Frequency'] = frequency  
        this_survey_df['Mandatory/Voluntary'] = mandatory_voluntary
        this_survey_df['Number of Reporting Units'] = num_units
        this_survey_df['Average Time to Complete (Per Form)'] = completion_time
        this_survey_df['Mailed Date'] = mailed_date
        this_survey_df['Due Date'] = due_date
        this_survey_df['Survey Information Page'] = info_page
        final_df = pd.concat([final_df,this_survey_df],ignore_index=True)
        #final_df = final_df.append(entry,ignore_index=True)
    final_df = final_df.fillna('')    
    return final_df


def create_address(unit):
    return unit[5] + '\n' + unit[6] + '\n' + unit[7] + ', ' + unit[8] + ' ' + unit[9]

height = 576 # height is usually this but may be modified in read_company_id_and_name

def read_company_id_and_name(calendar_path):  
    global height
    layout,dimensions = utils.get_page_layout(calendar_path)
    print(dimensions[1])
    height = dimensions[1]
    print(utils.get_page_layout(calendar_path))
    area_1 = '2,' + str(height - 36) + ',121,' + str(height-49)
    area_2 = '635,' + str(height - 33) + ',800,' + str(height-51)
    tables = camelot.read_pdf(calendar_path,pages='1',flavor='stream',table_areas=[area_1,area_2])
    df = tables[0].df
    
    company_id = df.iloc[0,1]
    df2 = tables[1].df
    company_name = df2.iloc[0,1].strip()
    
    return (company_id,company_name)


def get_survey_data(calendar_path,start_page,end_page):
    columns= ["""50,99,145,235,467,553,618.4,673.7,734.8,776"""]
    area = '-1,' + str(height - 96) + ',954,' + str(height - 463)
    tables = camelot.read_pdf(calendar_path,pages=str(start_page) + '-' + str(end_page),
                                   flavor='stream',table_areas=[area],columns=columns)
    dfs = []
    for table in tables:
        df = clean_survey_data(table.df)
        dfs.append(df)
    df = pd.concat(dfs)
    return df

def clean_survey_data(df):
    #df = df.iloc[8:] # dropping header
    df = merge_rows(df,[10])
    return df   

def merge_rows(df,to_add_new_line):
    new_df = pd.DataFrame(columns=df.columns)
    for index,row in df.iterrows():
        if row[0] and row[0].strip():
            new_df = new_df.append(row)
        else:
            for i in range(0,len(row)):
                if row[i] and row[i].strip():
                    whitespace = ' '
                    if i in to_add_new_line:
                        whitespace = '\n'
                    if not new_df.empty:
                        new_df.iat[-1,i] = new_df.iat[-1,i] + whitespace + row[i]
    return new_df


def get_first_entry_reporting_units(df):
    series = df[0]
    begin_entries_idx = None
    for index,value in series.iteritems():
        if value.isnumeric():
            begin_entries_idx = index
            break
    return begin_entries_idx    
    
    
def remove_junk_from_beginning(df):
    first_idx = get_first_entry_reporting_units(df)
    df = df.iloc[first_idx:]
    return df
    

def get_reporting_units(calendar_path,start_page):
    area = '-1,' + str(height - 96) + ',954,' + str(height - 463)
    tables = camelot.read_pdf(calendar_path,pages=str(start_page) + '-end',flavor='stream',
                              columns=["""50,93,140,211,252,301,
                                                  353,399,425,458,507,542,590,
                                                  642,684,724,846"""],table_areas=[area])
    
    dfs = []
    for table in tables:    
        df = clean_reporting_units_data(table.df)
        print(df.shape)
        print('^^ dimensions of reporting units page ^^')
        dfs.append(df)
    df = pd.concat(dfs)
    return df
        
    
def clean_reporting_units_data(df):
    df = merge_rows(df,[16,17])
    return df    

def extract_middle_page_survey(calendar_path,page_num):
    columns= ["""50,99,145,235,467,553,618.4,673.7,734.8,776"""]
    area = '-1,' + str(height - 96) + ',954,' + str(height - 463)
    tables = camelot.read_pdf(calendar_path,pages=str(page_num),
                                   flavor='stream',table_areas=[area],columns=columns)
    
    #tables= camelot.read_pdf(calendar_path,edge_tol=50000,pages=str(page_num),flavor='stream',columns=columns)
    #camelot.plot(tables[0],kind='text')
    #camelot.plot(tables[0],kind='contour')

    #plt.show()
    df = tables[0].df
    num_rows = df.shape[0]
    for position in range(num_rows):
        if '/' in df.iloc[position][1] and ':' in df.iloc[position][1]:
            df = df.iloc[:position] # what if position is 0?
            break
    df = clean_survey_data(df)
    return df
    
def extract_middle_page_units(calendar_path,page_num):
    columns=["""50,93,140,211,252,301,353,399,425,458,507,542,590,642,684,724,846"""]
    area = '-1,' + str(height - 96) + ',954,' + str(height - 463)
    tables = camelot.read_pdf(calendar_path,pages=str(page_num),
                                    flavor='stream',table_areas=[area],columns=columns)

    
    df = tables[0].df
    found_date = False
    found_unit_entry = False
    pattern = re.compile("^[0-9]{4,}$")
    for position in range(df.shape[0]):
        if found_date and pattern.match(df.iloc[position][0]):
            df = df.iloc[position:]
            found_unit_entry = True
            break
        if '/' in df.iloc[position][1] and ':' in df.iloc[position][1]:
            found_date = True
    if not found_unit_entry:
        df = pd.DataFrame()
    else:
        df = clean_reporting_units_data(df)
    return df
        

def get_survey_info_df(spreadsheet_path):
    df = pd.read_excel(spreadsheet_path,index_col=2)
    return df


def create_final_df(survey_df,reporting_units_df,survey_info):
    columns = ['Survey Name','Mandatory/Voluntary','Frequency','Mailed Date','Due Date','Response Date',
               'Company Contact','Average Time to Complete (Per Form)',
               'Survey Description','Survey Information Page',
               'Number of Reporting Units','Mailing Address']
    final_df = pd.DataFrame(columns=columns)
    survey_ids = survey_df[2].unique()
    
    survey_info_df = get_survey_info_df(survey_info) # survey info spreadsheet
    
    survey_df = survey_df.sort_values('3')
    for survey_id in survey_ids:
        forms = survey_df[survey_df[2] == survey_id]
        first = forms.iloc[0] # still need to go thru rest 
        entry = {}
        entry['Survey Name'] = first.iloc[3] + '\n' + '(' + survey_id + ')'
        entry['Frequency'] = first.iloc[7]
        
        id_for_survey_info = survey_id
        if (id_for_survey_info != 'M3'):
            id_for_survey_info = id_for_survey_info.rstrip('0123456789')
        try:
            mailed_date = survey_info_df.iloc[survey_info_df.index.get_loc(id_for_survey_info),4]
            if isinstance(mailed_date,datetime.datetime):
                mailed_date = mailed_date.date()
            entry['Mailed Date'] = mailed_date
            due_date = survey_info_df.iloc[survey_info_df.index.get_loc(id_for_survey_info),5]
            if isinstance(due_date,datetime.datetime):
                due_date = due_date.date()
            entry['Due Date'] = due_date
            entry['Mandatory/Voluntary'] = survey_info_df.iloc[survey_info_df.index.get_loc(id_for_survey_info),12] 
            entry['Average Time to Complete (Per Form)'] = survey_info_df.iloc[survey_info_df.index.get_loc(id_for_survey_info),18]
            entry['Survey Information Page'] = survey_info_df.iloc[survey_info_df.index.get_loc(id_for_survey_info),20]
        except KeyError:
            pass
        
        entry['Survey Description'] = first.iloc[4]
        num_units = 0
        addresses = ''
        company_contacts = ''
        response_dates = ''
        for index,form in forms.iterrows():
            form_name = form[0]
            num_units += int(form[1])
            units = reporting_units_df[(reporting_units_df[2] == form_name) & (reporting_units_df[1] == survey_id)]
            dashes = '----------------'
            for index,unit in units.iterrows():
                response_dates += unit[14] + '\n' + dashes + '\n'
                company_contacts += unit[16] + '\n' + dashes + '\n'
                addresses += 'Form: ' + form_name + '\n' + 'ID' + ' ' + unit[0] + '\n' + create_address(unit) + '\n' + dashes + '\n'
        response_dates = response_dates.strip().rstrip(dashes).rstrip()
        entry['Response Date'] = response_dates
        addresses = addresses.strip().rstrip(dashes).rstrip() # remove dashes and new lines at end
        entry['Mailing Address'] = addresses
        company_contacts = company_contacts.strip().rstrip(dashes).rstrip()
        entry['Company Contact'] = company_contacts
        entry['Number of Reporting Units'] = num_units 

        final_df = final_df.append(entry,ignore_index=True)
    final_df = final_df.fillna('')    
    return final_df
