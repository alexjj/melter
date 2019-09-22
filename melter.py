# -*- coding: utf-8 -*-

"""
1. open input sheet with all scenarios
inputs = pd.read_excel(file_name, sheet_name=None) #returns dict of all
2. Load attributes
2. For each df, process it - formatting and join with attributes
3. add new column df.Case = key 
4. Concat all dfs to one table
5. write to excel

"""

# %%
import pandas as pd
import numpy as np

from appJar import gui
from pathlib import Path



# %%
# input_file = '/home/alex/Dropbox/dev/Work/ipso/input-1/CasesfromSS.xlsx'
# output_file = '/home/alex/Dropbox/dev/Work/ipso/processed.csv'


#%%
def process_forecast(df, casename):
    """
    Takes a forecast dataframe, and returns the processed table:
    oil, water, steam, cyclic
    """
    if 'Data Check' in df.columns:
        df = df.drop('Data Check', 1)
    if 'Start Date' in df.columns:
        df = df.drop('Start Date', 1)
    # df = df[df.columns[:12*(years+1)]]    #trim number of years
    df['Case'] = casename
    df['Source/Sink'] = df['Source/Sink'].str.strip()  # Remove whitespace
    df['Source/Sink'] = df['Source/Sink'].str.upper()
    df = df.melt(id_vars=['Category', 'Project', 'Fluid', 'Units', 'Source/Sink', 'Case'], var_name='Date', value_name='Forecast')
    df = df.pivot_table(columns='Fluid', values='Forecast', index=['Category', 'Project', 'Source/Sink', 'Date', 'Case'], aggfunc=np.sum).reset_index().rename_axis(None, axis="columns")
    return df



# %%

def main(input_file, output_file):
    excel = pd.ExcelFile(input_file)
    frames = [process_forecast(excel.parse(f), f) for f in excel.sheet_names]
    all = pd.concat(frames, ignore_index=True)
    all.to_excel(output_file, sheet_name='input', index=False)
    
    if(app.questionBox("Complete!", "Forecasts Melted. Do you want to quit?")):
        app.stop()


# %%

def press(button):
    if button == "Melt":
        input_file = app.getEntry("input_file")
        output_file = app.getEntry("output_file")
        errors, error_msg = validate_inputs(input_file, output_file)
        if errors:
            app.errorBox("Error", "\n".join(error_msg), parent=None)
        else:
            main(input_file, output_file)
    else:
        app.stop()


def validate_inputs(input_file, output_file):
    errors = False
    error_msgs = []
    
    # Make sure xlsx is selected for input
    if Path(input_file).suffix.upper() != ".XLSX":
        errors = True
        error_msg.append("Please select an Excel input file")

    if len(output_file) < 1:
        errors = True
        error_msgs.append("Please select an output file")
    
    return(errors, error_msgs)


# Create GUI window
app = gui("Cymric IPSO Melter", useTtk=True)
app.setTtkTheme("default")
app.setSize(500, 200)
app.setFont(16)

# GUI interactions
app.addLabel("Choose forecast input spreadsheet:")
app.addFileEntry("input_file")

app.addLabel("Select output file:")
app.addFileEntry("output_file")

app.addButtons(["Melt", "Quit"], press)

app.go()