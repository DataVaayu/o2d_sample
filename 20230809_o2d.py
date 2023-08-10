# Tapping into the daily stock information

## Getting the credentials, reading the sheets and creating three dataframes

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = '1GTApxk_-pGgoPxXbhR0rSz_E9YuaB93kj8MjRzEuUOc'
O2D = 'O2D Client for Osaa Retail!A:N'


"""Shows basic usage of the Sheets API.
Prints values from a sample spreadsheet.
"""
creds = None
# The file token.json stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.json', 'w') as token:
        token.write(creds.to_json())

try:
    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API to get enquiry sheet
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=O2D).execute()
    values = result.get('values', [])
    


    if not values:
        print('No data found.')
        
    # creating the dataframe by reading enquiry sheet
    O2d_data = pd.DataFrame.from_records(values)   
    
    
except HttpError as err:
    print(err)

    
# __________________________________Sizing the dataframes_________________________________________________________________
    
# Setting the column names to the first rows that appear in the dataframes

dataframes=[O2d_data]

for i in dataframes:
    i.columns=i.loc[0]
    i.drop(0,axis=0,inplace=True)

### Creating a day, month and year column from the timestamp column
# creating the date column

O2d_data["Date"]=O2d_data["Timestamp"].apply(lambda x: x.split(" ")[0])

# filter out the dataframe such that there are no missing values in the date column
O2d_data = O2d_data[~O2d_data["Date"].isin([""])]

# setting the date column as a string column

O2d_data["Date"]=O2d_data["Date"].astype(str)

# Creating the day month and year column

O2d_data["Day"]=O2d_data["Date"].apply(lambda x: x.split("/")[0])
O2d_data["Month"]=O2d_data["Date"].apply(lambda x: x.split("/")[1])
O2d_data["Year"]=O2d_data["Date"].apply(lambda x: x.split("/")[2])

# concatenating the day month and year columns to create a Date2 column

for index, row in O2d_data.iterrows():
    O2d_data.loc[index,"Date2"]="-".join([row["Year"],row["Month"],row["Day"]])

# changing the Qty column to an integer column
O2d_data["Qty"] = O2d_data["Qty"].astype(int)

# ______________________________________Dash App Section ____________________________________________
### Creating a dash app for selection of dates and filtering the o2d data according to the entered dates

### Also creating a dashtable for hosting the data from the o2d client dataframe

from dash import Dash, callback, html, dcc, Input, Output, dash_table
import dash_bootstrap_components as dbc
from datetime import date, timedelta, datetime

app = Dash(__name__, external_stylesheets = [dbc.themes.LUX])

app.layout= dbc.Container([

    html.Br(),
    
    dbc.Row([   dbc.Col([

        dbc.Row([
            html.H3("Select Date Range",style={"background-color":"rgb(165,233,246)","text-align":"left"}),
            dcc.DatePickerRange(
                id="date-range",
                min_date_allowed=date(2023,1,1),
                initial_visible_month=date(2023,8,1),
                start_date=date(2023,8,1),
                end_date=date(2023,8,9),
                updatemode="singledate",
                display_format='MMM Do, YY',
                month_format='MMMM, YYYY',
                clearable=True,
                day_size=39,
                minimum_nights=1,
        ),

        ]),
        html.Br(),        
        dbc.Row([
            html.H3("Location-wise Sales",style={"background-color":"rgb(165,233,246)"}),
            dcc.Graph(id="location-sales")
        ]),
    ],width=4),

    html.Br(),
    html.Br(),
    
    dbc.Col([
    html.H3("Data from O2D client",style={"background-color":"rgb(165,233,246)","text-align":"center"}),
    dash_table.DataTable([{"name":i,"id":i} for i in O2d_data[["Pretture no.","Barcode no.","Customer Name","Department","Design no.","Qty","Colour","NOS"]].columns]
                         ,id="o2d-table",page_size=10,
                         style_data_conditional=[{
                             "if":{"filter_query":"{NOS}='Yes'",
                                   "column_id":"Design no."
                                   },
                             "background-color":"tomato",
                             "color":"white",
                         }],
                         ),

    ],width=8)
    ]),
    
    html.Br(),
    html.Br(),
    dbc.Row([
        dbc.Col([
            html.H3("Category-wise Distribution",style={"background-color":"rgb(165,233,246)","text-align":"center"}),
            dcc.Graph(id="category-graph"),
        ]),
    ]),

    html.Br(),
    html.Label(id="date-selection-show",style={"vertical-align":"middle"}),
    
])

@app.callback(
    Output("o2d-table","data"),
    Output("date-selection-show","children"),
    Output("location-sales","figure"),
    Output("category-graph","figure"),
    Input("date-range","start_date"),
    Input("date-range","end_date"),
)

def date_selection(start_date, end_date):
    print(f"start date: {datetime.strptime(start_date,'%Y-%m-%d')}")
    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)
    
    date_values=[]
    while start_date<=end_date:
        date_values.append(start_date.strftime("%Y-%m-%d"))
        start_date += timedelta(days=1)
        
    # Filter the o2d data based on the dates entered
    O2d_data2 = O2d_data[O2d_data["Date2"].isin(date_values)]

    # making a bar-chart of the Locations where products that were sold that have been bought

    O2d_data2_location  = O2d_data2.groupby("Department",as_index=False)["Qty"].sum()    
    fig2 = px.bar(O2d_data2_location,y="Qty",x="Department",height=250,width=500,text_auto=True)

    # making a bar chart for Category-wise sales during the time period
    O2d_data2_categories=O2d_data2.groupby("Category",as_index=False)["Qty"].sum()
    fig3 = px.bar(O2d_data2_categories,x="Category",y="Qty",text_auto=True)

    
        
    return O2d_data2[["Pretture no.","Barcode no.","Customer Name","Department","Design no.","Qty","Colour","NOS"]].to_dict("records"),f"The selected dates are : {date_values}",fig2, fig3  

if __name__=="__main__":
    app.run(debug=True,port=8021)