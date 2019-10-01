import pandas as pd
import numpy as np
from datetime import datetime

from report.template import template
import xlrd
import xlwt
from xlwt.Workbook import *
from pandas import ExcelWriter
import xlsxwriter
from selenium import webdriver
import time
import getpass
from pathlib import Path

loginusername =  (getpass.getuser())
iris= [
"CCED219",
"CCED217",
"CSA105",
"CSA104",
"CSA002",
"CSA001",
"CNI007",
"DTA023",
"DTA022",
"DTA021",
"DTA019",
"DTA015",
"DTA014",
"DTA012",
"DTA007",
"DTA005",
"DTC006",
"DCH055",
"DCH054",
"DCH1014",
"ELAN014",
"ELAN011",
"ELAN009",
"ELAN028",
"ELAN023",
"ELAN007",
"ERR008",
"FIJI007",
"FIJI010",
"FIJI014",
"FIJI081",
"INI109",
"INI009",
"INI007",
"INI003",
"INI002",
"INI000",
"SYS000"
"parseReported",
"PWR013",
"PWR053",
"PWR002",
"SEB BAIR Loader rejected",
"Listen port not responding",
"OVD004",
"Unable to perform overlapped",
"XMI007",
"Parse reported",
"PWR010"
]
Esc_manager =[]

print(''' Note:
    1. Alert and Ticket Report Names Should be given Perfectly with valid extentions(.xlsx or .csv)
    Eg:- Avaya alert.txt  -- Not Accepted
    Eg:-  Avayaalert.xlsx or Avayaalert.csv -- Accepted

    2. Place your report and image files on desktop with folder named shift_reports
    
    3. Please give req data (names of shift members, EM,reports and images) in this order don't change the order.
    present Shift Em ---- inline 1
    Next_Shift_Names of both team
    Next_Esc_manager 
    Ticket_Report_Name 
    Avaya_Alert_Name 
    Data_Alert_Name 
    carousel_images( gilabane and Carousel clients)

    save these details in a text file with ".txt" extension in shift_reports folder
''')
date_now  = datetime.now().strftime('%d-%m-%Y')
shift_time = datetime.now().time()
reports_path ="C:\\Users\\"+loginusername+"\Desktop\shift_reports\\"
filename = input("enter filename :")
dataFolder = Path(r'C:\Users\\'+loginusername+'\Desktop\\'+filename+".txt")
with open(dataFolder, 'r') as f:

    content = f.read().splitlines()
    ShiftEm =content[0]
    Next_Shift_Names = content[1]
    Next_Esc_manager = content [2]
    Ticket_Report_Name =content[3]
    Avaya_Alert_Name =content[4]
    Data_Alert_Name = content[5]
    gilbane= content[6]
    allclients= content[7]



def shifts(shift_time):
    global shift_name
    ES = "6:30"
    NS = "14:00"
    MS = "21:30"
    Evening_shift = datetime.strptime(ES, '%H:%M')
    Night_shift = datetime.strptime(NS, '%H:%M')
    Morning_shift = datetime.strptime(MS, '%H:%M')

    if shift_time > Evening_shift.time():
        shift_name ="ES"
    elif shift_time > Night_shift.time():
        shift_name ="NS"
    elif shift_time > Morning_shift.time():
        shift_name = "MS"
    return shift_name

def Shift_names(Ticket_Report_Name,reports_path):
    global report
    if 'xlsx' in Ticket_Report_Name:
        report = pd.read_excel(reports_path+Ticket_Report_Name, index_col=0)

    elif 'csv' in Ticket_Report_Name:
        report = pd.read_csv(reports_path+Ticket_Report_Name, index_col=0)

    else:
        print("Invalid Extension")
    report['Requester'] = report['Requester'].str.replace(" ", "")  # removes space from requester columns
    report = report.reset_index()
    names = pd.DataFrame(report['Requester']).drop_duplicates().values
    return names

def names(Shift_members):
    names = ["NaveenBommidi","NagaVenkataGogula","BalaTanuku","ganeshduppada","saigarimella","SarathGiduturi","BalaNaraharisetty",
             "P.Rkishore","RamatejaGaneshna","VenkataKishoreBaddukonda","RevanthMandava","SeshuYadla","adityajakkampudi"
             "KotanagaMotupalli","RajeshPattapagala","ShravanJuttiga","SohailSheik","VenkataSameerNamburi","SriDurgaSaiPrasadKatchetti",
             "DurgaprasadSirigineedi"]

    data_names=["KrishnaLokam","Saikalyan","AliShaik","maheshalladi","VenkataBathula","Sivalakkimsetty","NagaSannidhi","maheshalladi",
                "SekharSuthapalli","VaheedSheik"]
    Dataname =[]
    Avayaname =[]
    for i in data_names:
        for j in Shift_members:
            if str(j[0])==str(i):
                Dataname.append(j[0])

    for i in names:
        for j in Shift_members:
            if str(j[0])==str(i):
                Avayaname.append(j[0])

    return Avayaname,Dataname

def Data_Alert_Report(Data_Report_Name,shift,date_now,reports_path):
    global report
    if 'xlsx' in Data_Report_Name:
        report = pd.read_excel(reports_path+Data_Report_Name, index_col=0)

    elif 'csv' in Data_Report_Name:
        report = pd.read_csv(reports_path+Data_Report_Name, index_col=0)

    else:
        print("Invalid Extension")
    BCG = report[report['Client Name'] == 'BCG']
    On_Board_Alerts = report[report['Client Name'] != 'BCG']
    BCG_Count = len(BCG)
    On_Board_Count = len(On_Board_Alerts)
    writer = pd.ExcelWriter('Data_Alert_report '+"_ "+shift+"_"+ date_now+'.xlsx', engine='xlsxwriter')

    On_Board_Alerts.to_excel(writer, sheet_name='On_Board_Alerts')
    BCG.to_excel(writer, sheet_name='BCG')
    writer.save()
    return On_Board_Count,BCG_Count

def Avaya_Alert_Report(Avaya_Report_Name,shift,date_now,reports_path):
    global report  # Checks for Extension
    if 'xlsx' in Avaya_Report_Name:
        report = pd.read_excel(reports_path+Avaya_Report_Name, index_col=0)

    elif 'csv' in Avaya_Report_Name:
        report = pd.read_csv(reports_path+Avaya_Report_Name, index_col=0)

    else:
        print("Invalid Extension")

    # Seperates Alerts As per Source
    Alarm_traq = report[report['Source'] == 'AlarmTraq']
    Nectar = report[report['Source'] == 'Nectar']
    iris = report[report['Source'] == 'IRIS']

    #Gives Alert count for  each source
    Alarm_traq_count = len(Alarm_traq)
    Nectar_Count = len(Nectar)
    iris_Count = len(iris)

    # All alerts will be written to excel workbook
    writer = pd.ExcelWriter('Avaya_Alert_report '+"_ "+shift+"_"+ date_now+'.xlsx', engine='xlsxwriter')
    iris.to_excel(writer, sheet_name='IRIS')
    Alarm_traq.to_excel(writer, sheet_name='AlarmTrac')
    Nectar.to_excel(writer, sheet_name='Nectar')
    writer.save()
    # Returns alert Count details of all sources
    return Alarm_traq_count, Nectar_Count, iris_Count

def Avaya_Ticket_Report(Avaya_Ticket_Report,Employee,shift,date_now,reports_path):
    global report
    if 'xlsx' in Avaya_Ticket_Report:
        report = pd.read_excel(reports_path+Avaya_Ticket_Report, index_col=0)

    elif 'csv' in Avaya_Ticket_Report:
        report = pd.read_csv(reports_path+Avaya_Ticket_Report, index_col=0)

    else:
        print("Invalid Extension")
    report['Requester'] = report['Requester'].str.replace(" ","")  # removes space from requester columns
    #report['Requester']= report['Requester'].str.lower() # converts names to lowercase
    df = pd.DataFrame(report)

    # tickets = df.query('Requester in("Venkata Kishore Baddukonda","Bala Naraharisetty","P.R kishore","sai garimella" Ramateja Ganeshna)')
    tickets = df.query('Requester in(' + str(Employee) + ')') # Gives Tickets generated by shift members
    tickets['Target'] = np.where(tickets['Priority'] =="High" ,'00:25:00',
                                np.where(tickets['Priority'] == "Medium",'00:30:00',
                                 np.where(tickets['Priority'] == "Normal",'00:30:00',
                                  np.where(tickets['Priority']== "Low",'00:30:00',
                                   np.where(tickets['Priority'] == "Critical", '00:10:00',tickets['Priority'])))))
    tickets['First Response Time (HH:MM:SS)'] = tickets['First Response Time (HH:MM:SS)'].replace({0: '00:00:00'})
    tickets['First Response Time (HH:MM:SS)'] = pd.to_datetime(tickets['First Response Time (HH:MM:SS)'], format='%H:%M:%S').dt.time
    tickets['Target'] = pd.to_datetime(tickets['Target'], format='%H:%M:%S').dt.time
    tickets['SLA'] = np.where(tickets['First Response Time (HH:MM:SS)']< tickets['Target'],"INSLA","OSLA")

    breaches = tickets[tickets['SLA'] == "OSLA"]
    iris_tickets = tickets[tickets['Subject'].str.contains('|'.join(iris))] # Iris Tickets

    total_tickets  = tickets[~tickets['Subject'].str.contains('|'.join(iris))] # dataframe with out IRIS tickets

    at_tickets = total_tickets[total_tickets.Subject.str.contains(",")] # alarm trac tickets

    nc_tickets = total_tickets[total_tickets.Subject.str.contains(",")== False] # Nectar Tickets
    # for Ticket Length
    iris_tickets_len = len(iris_tickets)
    AlarmTrac_Tkt_len = len(at_tickets)
    Nectar_Tkt_Len = len(nc_tickets)

    #For Priority wise Details with count of AlarmTrac
    At_Critical_P1_INSLA = len(at_tickets.loc[(at_tickets['Priority']=='Critical') &(at_tickets['SLA']=="INSLA")])
    At_Critical_P1_OSLA = len(at_tickets.loc[(at_tickets['Priority'] == 'Critical') & (at_tickets['SLA'] == "OSLA")])

    At_High_P2_INSLA = len(at_tickets.loc[(at_tickets['Priority'] == 'High') & (at_tickets['SLA'] == "INSLA")])
    At_High_P2_OSLA = len(at_tickets.loc[(at_tickets['Priority'] == 'High') & (at_tickets['SLA'] == "OSLA")])

    At_Medium_P3_INSLA = len(at_tickets.loc[(at_tickets['Priority'] == 'Medium') & (at_tickets['SLA'] == "INSLA")])
    At_Medium_P3_OSLA = len(at_tickets.loc[(at_tickets['Priority'] == 'Medium') & (at_tickets['SLA'] == "OSLA")])

    At_Normal_P4_INSLA = len(at_tickets.loc[(at_tickets['Priority'] == 'Normal') & (at_tickets['SLA'] == "INSLA")])
    At_Low_P4_INSLA = len(at_tickets.loc[(at_tickets['Priority'] == 'Low') & (at_tickets['SLA'] == "INSLA")])

    At_Normal_P4_OSLA = len(at_tickets.loc[(at_tickets['Priority'] == 'Normal') & (at_tickets['SLA'] == "OSLA")])
    At_Low_P4_OSLA = len(at_tickets.loc[(at_tickets['Priority'] == 'Low') & (at_tickets['SLA'] == "OSLA")])

    At_Normal_INSLA_Count =At_Normal_P4_INSLA + At_Low_P4_INSLA
    At_Low_OSLA_Count = At_Normal_P4_OSLA + At_Low_P4_OSLA
    AlarmTrac_Priorities_Count = [At_Critical_P1_INSLA, At_Critical_P1_OSLA,
                            At_High_P2_INSLA, At_High_P2_OSLA, At_Medium_P3_INSLA,
                            At_Medium_P3_OSLA, At_Normal_INSLA_Count, At_Low_OSLA_Count]

    # For Priority wise Details with count of Nectar
    Nc_Critical_P1_INSLA = len(nc_tickets.loc[(nc_tickets['Priority'] == 'Critical') & (nc_tickets['SLA'] == "INSLA")])
    Nc_Critical_P1_OSLA = len(nc_tickets.loc[(nc_tickets['Priority'] == 'Critical') & (nc_tickets['SLA'] == "OSLA")])

    Nc_High_P2_INSLA = len(nc_tickets.loc[(nc_tickets['Priority'] == 'High') & (nc_tickets['SLA'] == "INSLA")])
    Nc_High_P2_OSLA = len(nc_tickets.loc[(nc_tickets['Priority'] == 'High') & (nc_tickets['SLA'] == "OSLA")])

    Nc_Medium_P3_INSLA = len(nc_tickets.loc[(nc_tickets['Priority'] == 'Medium') & (nc_tickets['SLA'] == "INSLA")])
    Nc_Medium_P3_OSLA = len(nc_tickets.loc[(nc_tickets['Priority'] == 'Medium') & (nc_tickets['SLA'] == "OSLA")])

    Nc_Normal_P4_INSLA = len(nc_tickets.loc[(nc_tickets['Priority'] == 'Normal') & (nc_tickets['SLA'] == "INSLA")])
    Nc_Low_P4_INSLA = len(nc_tickets.loc[(nc_tickets['Priority'] == 'Low') & (nc_tickets['SLA'] == "INSLA")])

    Nc_Normal_P4_OSLA = len(nc_tickets.loc[(nc_tickets['Priority'] == 'Normal') & (nc_tickets['SLA'] == "OSLA")])
    Nc_Low_P4_OSLA = len(nc_tickets.loc[(nc_tickets['Priority'] == 'Low') & (nc_tickets['SLA'] == "OSLA")])

    Nc_Normal_INSLA_Count = Nc_Normal_P4_INSLA + Nc_Low_P4_INSLA
    Nc_Low_OSLA_Count = Nc_Normal_P4_OSLA + Nc_Low_P4_OSLA

    Nectar_Priorities_Count = [Nc_Critical_P1_INSLA, Nc_Critical_P1_OSLA,
                            Nc_High_P2_INSLA, Nc_High_P2_OSLA, Nc_Medium_P3_INSLA,
                            Nc_Medium_P3_OSLA, Nc_Normal_INSLA_Count, Nc_Low_OSLA_Count]

    # For Priority wise Details with count of IRIS
    iris_Critical_P1_INSLA = len(iris_tickets.loc[(iris_tickets['Priority'] == 'Critical') & (iris_tickets['SLA'] == "INSLA")])
    iris_Critical_P1_OSLA = len(iris_tickets.loc[(iris_tickets['Priority'] == 'Critical') & (iris_tickets['SLA'] == "OSLA")])

    iris_High_P2_INSLA = len(iris_tickets.loc[(iris_tickets['Priority'] == 'High') & (iris_tickets['SLA'] == "INSLA")])
    iris_High_P2_OSLA = len(iris_tickets.loc[(iris_tickets['Priority'] == 'High') & (iris_tickets['SLA'] == "OSLA")])

    iris_Medium_P3_INSLA = len(iris_tickets.loc[(iris_tickets['Priority'] == 'Medium') & (iris_tickets['SLA'] == "INSLA")])
    iris_Medium_P3_OSLA = len(iris_tickets.loc[(iris_tickets['Priority'] == 'Medium') & (iris_tickets['SLA'] == "OSLA")])

    iris_Normal_P4_INSLA = len(iris_tickets.loc[(iris_tickets['Priority'] == 'Normal') & (iris_tickets['SLA'] == "INSLA")])
    iris_Low_P4_INSLA = len(iris_tickets.loc[(iris_tickets['Priority'] == 'Low') & (iris_tickets['SLA'] == "INSLA")])

    iris_Normal_P4_OSLA = len(iris_tickets.loc[(iris_tickets['Priority'] == 'Normal') & (iris_tickets['SLA'] == "OSLA")])
    iris_Low_P4_OSLA = len(iris_tickets.loc[(iris_tickets['Priority'] == 'Low') & (iris_tickets['SLA'] == "OSLA")])

    iris_Normal_INSLA_Count = iris_Normal_P4_INSLA + iris_Low_P4_INSLA
    iris_Low_OSLA_Count = iris_Normal_P4_OSLA + iris_Low_P4_OSLA
    Iris_Priorities_Count = [iris_Critical_P1_INSLA, iris_Critical_P1_OSLA,
                               iris_High_P2_INSLA, iris_High_P2_OSLA, iris_Medium_P3_INSLA,
                               iris_Medium_P3_OSLA, iris_Normal_INSLA_Count, iris_Low_OSLA_Count]
      # df.loc[(df['column_name'] >= A) & (df['column_name'] <= B)]
    # print("," .join(column_names) )
    # tickets.to_csv('tickets.csv', encoding='utf-8')
    writer = pd.ExcelWriter('Avaya_Ticket_report'+ '_ '+shift+'_'+ date_now+'.xlsx', engine='xlsxwriter')
    iris_tickets.to_excel(writer, sheet_name='IRIS')
    at_tickets.to_excel(writer, sheet_name='AlarmTrac')
    nc_tickets.to_excel(writer, sheet_name='Nectar')
    writer.save()

    return Nectar_Tkt_Len,AlarmTrac_Tkt_len,iris_tickets_len,AlarmTrac_Priorities_Count,Nectar_Priorities_Count,Iris_Priorities_Count,breaches

def Data_Ticket_Report(Data_Ticket_Report,Employee,shift,date_now,reports_path):
    global report
    if 'xlsx' in Data_Ticket_Report:
        report = pd.read_excel(reports_path+Data_Ticket_Report, index_col=0)

    elif 'csv' in Data_Ticket_Report:
        report = pd.read_csv(reports_path+Data_Ticket_Report, index_col=0)

    else:
        print("Invalid Extension")

    report['Requester'] = report['Requester'].str.replace(" ", "")  # removes space from requester columns
    #report['Requester'] = report['Requester'].str.lower()  # converts names to lowercase
    df = pd.DataFrame(report)

    # tickets = df.query('Requester in("Venkata Kishore Baddukonda","Bala Naraharisetty","P.R kishore","sai garimella" Ramateja Ganeshna)')
    Data_tickets = df.query('Requester in(' + str(Employee) + ')')  # Gives Tickets generated by shift members

    BCG_tickets = Data_tickets[Data_tickets['Client Name']=="BCG"]  # BCG tickets
    BCG_tickets['Target'] = np.where(BCG_tickets['Priority'] == "High", '00:15:00',
                                     np.where(BCG_tickets['Priority'] == "Medium", '00:20:00',
                                              np.where(BCG_tickets['Priority'] == "Normal", '00:20:00',
                                                       np.where(BCG_tickets['Priority'] == "Low", '00:20:00',
                                                                np.where(BCG_tickets['Priority'] == "Critical",
                                                                         '00:10:00',
                                                                         BCG_tickets['Priority'])))))
    BCG_tickets['First Response Time (HH:MM:SS)'] = BCG_tickets['First Response Time (HH:MM:SS)'].replace(
        {0: '00:00:00'})
    BCG_tickets['First Response Time (HH:MM:SS)'] = pd.to_datetime(BCG_tickets['First Response Time (HH:MM:SS)'],
                                                                   format='%H:%M:%S').dt.time
    BCG_tickets['Target'] = pd.to_datetime(BCG_tickets['Target'], format='%H:%M:%S').dt.time
    BCG_tickets['SLA'] = np.where(BCG_tickets['First Response Time (HH:MM:SS)'] < BCG_tickets['Target'], "INSLA",
                                  "OSLA")
    Bcg_tkt_count = len(BCG_tickets)
    BCG_Breaches = BCG_tickets[BCG_tickets['SLA'] == "OSLA"]
    # For Priority wise Details with count of BCG
    BCG_Critical_P1_INSLA = len(
        BCG_tickets.loc[(BCG_tickets['Priority'] == 'Critical') & (BCG_tickets['SLA'] == "INSLA")])
    BCG_Critical_P1_OSLA = len(
        BCG_tickets.loc[(BCG_tickets['Priority'] == 'Critical') & (BCG_tickets['SLA'] == "OSLA")])

    BCG_High_P2_INSLA = len(BCG_tickets.loc[(BCG_tickets['Priority'] == 'High') & (BCG_tickets['SLA'] == "INSLA")])
    BCG_High_P2_OSLA = len(BCG_tickets.loc[(BCG_tickets['Priority'] == 'High') & (BCG_tickets['SLA'] == "OSLA")])

    BCG_Medium_P3_INSLA = len(BCG_tickets.loc[(BCG_tickets['Priority'] == 'Medium') & (BCG_tickets['SLA'] == "INSLA")])
    BCG_Medium_P3_OSLA = len(BCG_tickets.loc[(BCG_tickets['Priority'] == 'Medium') & (BCG_tickets['SLA'] == "OSLA")])

    BCG_Normal_P4_INSLA = len(BCG_tickets.loc[(BCG_tickets['Priority'] == 'Normal') & (BCG_tickets['SLA'] == "INSLA")])
    BCG_Low_P4_INSLA = len(BCG_tickets.loc[(BCG_tickets['Priority'] == 'Low') & (BCG_tickets['SLA'] == "INSLA")])

    BCG_Normal_P4_OSLA = len(BCG_tickets.loc[(BCG_tickets['Priority'] == 'Normal') & (BCG_tickets['SLA'] == "OSLA")])
    BCG_Low_P4_OSLA = len(BCG_tickets.loc[(BCG_tickets['Priority'] == 'Low') & (BCG_tickets['SLA'] == "OSLA")])

    BCG_Normal_INSLA_Count = BCG_Normal_P4_INSLA + BCG_Low_P4_INSLA
    BCG_Low_OSLA_Count = BCG_Normal_P4_OSLA + BCG_Low_P4_OSLA

    BCG_Priorities_Count = [BCG_Critical_P1_INSLA, BCG_Critical_P1_OSLA,
                            BCG_High_P2_INSLA, BCG_High_P2_OSLA, BCG_Medium_P3_INSLA,
                            BCG_Medium_P3_OSLA, BCG_Normal_INSLA_Count, BCG_Low_OSLA_Count]


    OnBoard_tickets = Data_tickets[Data_tickets['Client Name'] !="BCG"]  # OnBoard Tickets
    OnBoard_tickets['Target'] = np.where(OnBoard_tickets['Priority'] == "High", '00:25:00',
                                         np.where(OnBoard_tickets['Priority'] == "Medium", '00:30:00',
                                                  np.where(OnBoard_tickets['Priority'] == "Normal", '00:30:00',
                                                           np.where(OnBoard_tickets['Priority'] == "Low", '00:30:00',
                                                                    np.where(OnBoard_tickets['Priority'] == "Critical",
                                                                             '00:10:00',
                                                                             OnBoard_tickets['Priority'])))))
    OnBoard_tickets['First Response Time (HH:MM:SS)'] = OnBoard_tickets['First Response Time (HH:MM:SS)'].replace(
        {0: '00:00:00'})
    OnBoard_tickets['First Response Time (HH:MM:SS)'] = pd.to_datetime(
        OnBoard_tickets['First Response Time (HH:MM:SS)'],
        format='%H:%M:%S').dt.time
    OnBoard_tickets['Target'] = pd.to_datetime(OnBoard_tickets['Target'], format='%H:%M:%S').dt.time
    OnBoard_tickets['SLA'] = np.where(OnBoard_tickets['First Response Time (HH:MM:SS)'] < OnBoard_tickets['Target'],
                                      "INSLA", "OSLA")

    OnBoard_tkt_count = len(OnBoard_tickets)
    OnBoard_Breaches = OnBoard_tickets[OnBoard_tickets['SLA'] == "OSLA"]

    # For Priority wise Details with count of OnBoard
    OnBoard_Critical_P1_INSLA = len(
        OnBoard_tickets.loc[(OnBoard_tickets['Priority'] == 'Critical') & (OnBoard_tickets['SLA'] == "INSLA")])
    OnBoard_Critical_P1_OSLA = len(
        OnBoard_tickets.loc[(OnBoard_tickets['Priority'] == 'Critical') & (OnBoard_tickets['SLA'] == "OSLA")])

    OnBoard_High_P2_INSLA = len(
        OnBoard_tickets.loc[(OnBoard_tickets['Priority'] == 'High') & (OnBoard_tickets['SLA'] == "INSLA")])
    OnBoard_High_P2_OSLA = len(
        OnBoard_tickets.loc[(OnBoard_tickets['Priority'] == 'High') & (OnBoard_tickets['SLA'] == "OSLA")])

    OnBoard_Medium_P3_INSLA = len(
        OnBoard_tickets.loc[(OnBoard_tickets['Priority'] == 'Medium') & (OnBoard_tickets['SLA'] == "INSLA")])
    OnBoard_Medium_P3_OSLA = len(
        OnBoard_tickets.loc[(OnBoard_tickets['Priority'] == 'Medium') & (OnBoard_tickets['SLA'] == "OSLA")])

    OnBoard_Normal_P4_INSLA = len(
        OnBoard_tickets.loc[(OnBoard_tickets['Priority'] == 'Normal') & (OnBoard_tickets['SLA'] == "INSLA")])
    OnBoard_Low_P4_INSLA = len(
        OnBoard_tickets.loc[(OnBoard_tickets['Priority'] == 'Low') & (OnBoard_tickets['SLA'] == "INSLA")])

    OnBoard_Normal_P4_OSLA = len(
        OnBoard_tickets.loc[(OnBoard_tickets['Priority'] == 'Normal') & (OnBoard_tickets['SLA'] == "OSLA")])
    OnBoard_Low_P4_OSLA = len(
        OnBoard_tickets.loc[(OnBoard_tickets['Priority'] == 'Low') & (OnBoard_tickets['SLA'] == "OSLA")])

    OnBoard_Normal_INSLA_Count = OnBoard_Normal_P4_INSLA + OnBoard_Low_P4_INSLA
    OnBoard_Low_OSLA_Count = OnBoard_Normal_P4_OSLA + OnBoard_Low_P4_OSLA

    OnBoard_Priorities_Count = [OnBoard_Critical_P1_INSLA, OnBoard_Critical_P1_OSLA,
                                OnBoard_High_P2_INSLA, OnBoard_High_P2_OSLA, OnBoard_Medium_P3_INSLA,
                                OnBoard_Medium_P3_OSLA, OnBoard_Normal_INSLA_Count, OnBoard_Low_OSLA_Count]


    writer = pd.ExcelWriter('Data_Ticket_report'+ '_ '+shift+'_'+ date_now+'.xlsx', engine='xlsxwriter')
    BCG_tickets.to_excel(writer, sheet_name='BCG')
    OnBoard_tickets.to_excel(writer, sheet_name='OB_Ticket')
    writer.save()

    return Bcg_tkt_count, OnBoard_tkt_count,BCG_Priorities_Count,OnBoard_Priorities_Count,BCG_Breaches,OnBoard_Breaches

def sla_dashboard(Avaya_Tkt_count_And_Proirities, Data_Tkt_count_And_Proirities):
    Nectar_Total = (Avaya_Tkt_count_And_Proirities[4][0] + Avaya_Tkt_count_And_Proirities[4][2] +
                    Avaya_Tkt_count_And_Proirities[4][4]
                    + Avaya_Tkt_count_And_Proirities[4][6] + Avaya_Tkt_count_And_Proirities[4][1] +
                    Avaya_Tkt_count_And_Proirities[4][3] +
                    Avaya_Tkt_count_And_Proirities[4][5]
                    + Avaya_Tkt_count_And_Proirities[4][7])
    AlarmTrac_Total = (Avaya_Tkt_count_And_Proirities[3][0] + Avaya_Tkt_count_And_Proirities[3][2]
                       + Avaya_Tkt_count_And_Proirities[3][4] + Avaya_Tkt_count_And_Proirities[3][6] +
                       Avaya_Tkt_count_And_Proirities[3][1] + Avaya_Tkt_count_And_Proirities[3][3] +
                       Avaya_Tkt_count_And_Proirities[3][5] + Avaya_Tkt_count_And_Proirities[3][7])

    Iris_Total = (Avaya_Tkt_count_And_Proirities[5][0] + Avaya_Tkt_count_And_Proirities[5][2] +
                  Avaya_Tkt_count_And_Proirities[5][4]
                  + Avaya_Tkt_count_And_Proirities[5][6] + Avaya_Tkt_count_And_Proirities[5][1] +
                  Avaya_Tkt_count_And_Proirities[5][3]
                  + Avaya_Tkt_count_And_Proirities[5][5] + Avaya_Tkt_count_And_Proirities[5][7])
    OnBoard_Total = (Data_Tkt_count_And_Proirities[3][0] + Data_Tkt_count_And_Proirities[3][2] +
                     Data_Tkt_count_And_Proirities[3][4]
                     + Data_Tkt_count_And_Proirities[3][6] + Data_Tkt_count_And_Proirities[3][1] +
                     Data_Tkt_count_And_Proirities[3][3]
                     + Data_Tkt_count_And_Proirities[3][5] + Data_Tkt_count_And_Proirities[3][7])
    BCG_Total = (Data_Tkt_count_And_Proirities[2][0] + Data_Tkt_count_And_Proirities[2][2] +
                 Data_Tkt_count_And_Proirities[2][4]
                 + Data_Tkt_count_And_Proirities[2][6] + Data_Tkt_count_And_Proirities[2][1] +
                 Data_Tkt_count_And_Proirities[2][3] +
                 Data_Tkt_count_And_Proirities[2][5]
                 + Data_Tkt_count_And_Proirities[2][7])
    return Nectar_Total, AlarmTrac_Total, Iris_Total, OnBoard_Total, BCG_Total


def grand_total(Avaya_Tkt_count_And_Proirities, Data_Tkt_count_And_Proirities):
    P1_grand_total = (Avaya_Tkt_count_And_Proirities[4][0] + Avaya_Tkt_count_And_Proirities[4][1] +
                      Avaya_Tkt_count_And_Proirities[3][0] + Avaya_Tkt_count_And_Proirities[3][1] +
                      Avaya_Tkt_count_And_Proirities[5][0] + Avaya_Tkt_count_And_Proirities[5][1] +
                      Data_Tkt_count_And_Proirities[2][0] + Data_Tkt_count_And_Proirities[2][1] +
                      Data_Tkt_count_And_Proirities[3][0] + Data_Tkt_count_And_Proirities[3][1])
    P2_grand_total = (Avaya_Tkt_count_And_Proirities[4][2] + Avaya_Tkt_count_And_Proirities[4][3] +
                      Avaya_Tkt_count_And_Proirities[3][2] + Avaya_Tkt_count_And_Proirities[3][3] +
                      Avaya_Tkt_count_And_Proirities[5][2] + Avaya_Tkt_count_And_Proirities[5][3] +
                      Data_Tkt_count_And_Proirities[3][2] + Data_Tkt_count_And_Proirities[3][3] +
                      Data_Tkt_count_And_Proirities[2][2] + Data_Tkt_count_And_Proirities[2][3])
    P3_grand_total = (Avaya_Tkt_count_And_Proirities[4][4] + Avaya_Tkt_count_And_Proirities[4][5] +
                      Avaya_Tkt_count_And_Proirities[3][4] + Avaya_Tkt_count_And_Proirities[3][5] +
                      Avaya_Tkt_count_And_Proirities[5][4] + Avaya_Tkt_count_And_Proirities[5][5] +
                      Data_Tkt_count_And_Proirities[3][4] + Data_Tkt_count_And_Proirities[3][5] +
                      Data_Tkt_count_And_Proirities[2][4] + Data_Tkt_count_And_Proirities[2][5])
    P4_grand_total = (Avaya_Tkt_count_And_Proirities[4][6] + Avaya_Tkt_count_And_Proirities[4][7] +
                      Avaya_Tkt_count_And_Proirities[3][6] + Avaya_Tkt_count_And_Proirities[3][7] +
                      Avaya_Tkt_count_And_Proirities[5][6] + Avaya_Tkt_count_And_Proirities[5][7] +
                      Data_Tkt_count_And_Proirities[3][6] + Data_Tkt_count_And_Proirities[3][7] +
                      Data_Tkt_count_And_Proirities[2][6] + Data_Tkt_count_And_Proirities[2][7])
    return P1_grand_total, P2_grand_total, P3_grand_total, P4_grand_total


def grand_total_sla(Avaya_Tkt_count_And_Proirities, Data_Tkt_count_And_Proirities):
    grand_INSLA = (Avaya_Tkt_count_And_Proirities[4][0] + Avaya_Tkt_count_And_Proirities[4][2] +
                   Avaya_Tkt_count_And_Proirities[4][4] + Avaya_Tkt_count_And_Proirities[4][6] +
                   Avaya_Tkt_count_And_Proirities[3][0] + Avaya_Tkt_count_And_Proirities[3][2] +
                   Avaya_Tkt_count_And_Proirities[3][4] + Avaya_Tkt_count_And_Proirities[3][6] +
                   Avaya_Tkt_count_And_Proirities[5][0] + Avaya_Tkt_count_And_Proirities[5][2] +
                   Avaya_Tkt_count_And_Proirities[5][4] + Avaya_Tkt_count_And_Proirities[5][6] +
                   Data_Tkt_count_And_Proirities[3][0] + Data_Tkt_count_And_Proirities[3][2] +
                   Data_Tkt_count_And_Proirities[3][4] + Data_Tkt_count_And_Proirities[3][6] +
                   Data_Tkt_count_And_Proirities[2][0] + Data_Tkt_count_And_Proirities[2][2] +
                   Data_Tkt_count_And_Proirities[2][4] + Data_Tkt_count_And_Proirities[2][6])
    grand_OSLA = (Avaya_Tkt_count_And_Proirities[4][1] + Avaya_Tkt_count_And_Proirities[4][3] +
                  Avaya_Tkt_count_And_Proirities[4][5] + Avaya_Tkt_count_And_Proirities[4][7] +
                  Avaya_Tkt_count_And_Proirities[3][1] + Avaya_Tkt_count_And_Proirities[3][3] +
                  Avaya_Tkt_count_And_Proirities[3][5] + Avaya_Tkt_count_And_Proirities[3][7] +
                  Avaya_Tkt_count_And_Proirities[5][1] + Avaya_Tkt_count_And_Proirities[5][3] +
                  Avaya_Tkt_count_And_Proirities[5][5] + Avaya_Tkt_count_And_Proirities[5][7] +
                  Data_Tkt_count_And_Proirities[3][1] + Data_Tkt_count_And_Proirities[3][3] +
                  Data_Tkt_count_And_Proirities[3][5] + Data_Tkt_count_And_Proirities[3][7] +
                  Data_Tkt_count_And_Proirities[2][1] + Data_Tkt_count_And_Proirities[2][3] +
                  Data_Tkt_count_And_Proirities[2][5] + Data_Tkt_count_And_Proirities[2][7])
    return grand_INSLA, grand_OSLA


def sla_dashboard_percentage(Avaya_Tkt_count_And_Proirities, Data_Tkt_count_And_Proirities, Total_ticket_count):


    Nectar_dashboard_p1 = calculatePercentage(Avaya_Tkt_count_And_Proirities[4][0], (Avaya_Tkt_count_And_Proirities[4][0] + Avaya_Tkt_count_And_Proirities[4][1]))

    Nectar_dashboard_p2 = calculatePercentage(Avaya_Tkt_count_And_Proirities[4][2], (Avaya_Tkt_count_And_Proirities[4][2] + Avaya_Tkt_count_And_Proirities[4][3]))

    Nectar_dashboard_p3 = calculatePercentage(Avaya_Tkt_count_And_Proirities[4][4], (Avaya_Tkt_count_And_Proirities[4][4] + Avaya_Tkt_count_And_Proirities[4][5]))

    Nectar_dashboard_p4 = calculatePercentage(Avaya_Tkt_count_And_Proirities[4][6], (Avaya_Tkt_count_And_Proirities[4][6] + Avaya_Tkt_count_And_Proirities[4][7]))



    Alarmtrac_dashboard_p1 = calculatePercentage(Avaya_Tkt_count_And_Proirities[3][0], (Avaya_Tkt_count_And_Proirities[3][0] + Avaya_Tkt_count_And_Proirities[3][1]))

    Alarmtrac_dashboard_p2 = calculatePercentage(Avaya_Tkt_count_And_Proirities[3][2], (Avaya_Tkt_count_And_Proirities[3][2] + Avaya_Tkt_count_And_Proirities[3][3]))

    Alarmtrac_dashboard_p3 = calculatePercentage(Avaya_Tkt_count_And_Proirities[3][4], (Avaya_Tkt_count_And_Proirities[3][4] + Avaya_Tkt_count_And_Proirities[3][5]))

    Alarmtrac_dashboard_p4 = calculatePercentage(Avaya_Tkt_count_And_Proirities[3][6], (Avaya_Tkt_count_And_Proirities[3][6] + Avaya_Tkt_count_And_Proirities[3][7]))



    Iris_dashboard_p1 = calculatePercentage(Avaya_Tkt_count_And_Proirities[5][0], (Avaya_Tkt_count_And_Proirities[5][0] + Avaya_Tkt_count_And_Proirities[5][1]))

    Iris_dashboard_p2 = calculatePercentage(Avaya_Tkt_count_And_Proirities[5][2], (Avaya_Tkt_count_And_Proirities[5][2] + Avaya_Tkt_count_And_Proirities[5][3]))

    Iris_dashboard_p3 = calculatePercentage(Avaya_Tkt_count_And_Proirities[5][4], (Avaya_Tkt_count_And_Proirities[5][4] + Avaya_Tkt_count_And_Proirities[5][5]))

    Iris_dashboard_p4 = calculatePercentage(Avaya_Tkt_count_And_Proirities[5][6], (Avaya_Tkt_count_And_Proirities[5][6] + Avaya_Tkt_count_And_Proirities[5][7]))



    OnBoard_dashboard_p1 = calculatePercentage(Data_Tkt_count_And_Proirities[3][0], (Data_Tkt_count_And_Proirities[3][0] + Data_Tkt_count_And_Proirities[3][1]))

    OnBoard_dashboard_p2 = calculatePercentage(Data_Tkt_count_And_Proirities[3][2], (Data_Tkt_count_And_Proirities[3][2] + Data_Tkt_count_And_Proirities[3][3]))

    OnBoard_dashboard_p3 = calculatePercentage(Data_Tkt_count_And_Proirities[3][4], (Data_Tkt_count_And_Proirities[3][4] + Data_Tkt_count_And_Proirities[3][5]))

    OnBoard_dashboard_p4 = calculatePercentage(Data_Tkt_count_And_Proirities[3][6], (Data_Tkt_count_And_Proirities[3][6] + Data_Tkt_count_And_Proirities[3][7]))


    BCG__dashboard_p1 = calculatePercentage(Data_Tkt_count_And_Proirities[2][0], ( Data_Tkt_count_And_Proirities[2][0]+ Data_Tkt_count_And_Proirities[2][1]))

    BCG__dashboard_p2 = calculatePercentage(Data_Tkt_count_And_Proirities[2][2], (Data_Tkt_count_And_Proirities[2][2] + Data_Tkt_count_And_Proirities[2][3]))

    BCG__dashboard_p3 = calculatePercentage(Data_Tkt_count_And_Proirities[2][4], (Data_Tkt_count_And_Proirities[2][4] + Data_Tkt_count_And_Proirities[2][5]))

    BCG__dashboard_p4 = calculatePercentage(Data_Tkt_count_And_Proirities[2][6], (Data_Tkt_count_And_Proirities[2][6] + Data_Tkt_count_And_Proirities[2][7]))

    Nectar_dashboard_values = [Nectar_dashboard_p1, Nectar_dashboard_p2, Nectar_dashboard_p3, Nectar_dashboard_p4]
    Alarmtrac_dashboard_values = [Alarmtrac_dashboard_p1, Alarmtrac_dashboard_p2, Alarmtrac_dashboard_p3,Alarmtrac_dashboard_p4]
    OnBoard_dashboard_values = [OnBoard_dashboard_p1, OnBoard_dashboard_p2, OnBoard_dashboard_p3, OnBoard_dashboard_p4]
    Iris_dashboard_values = [Iris_dashboard_p1, Iris_dashboard_p2, Iris_dashboard_p3, Iris_dashboard_p4]
    BCG__dashboard_Values = [BCG__dashboard_p1, BCG__dashboard_p2, BCG__dashboard_p3, BCG__dashboard_p4]
    return Nectar_dashboard_values, Alarmtrac_dashboard_values, Iris_dashboard_values, OnBoard_dashboard_values, BCG__dashboard_Values

def calculatePercentage(dividend, divisor):
    percentage = 0 if divisor == 0 else round(((dividend/divisor) * 100))
    return percentage

def Breaches(Avaya_Tkt_count_And_Proirities,Data_Tkt_count_And_Proirities):
    BCG = pd.DataFrame(Data_Tkt_count_And_Proirities[4])
    On_Board =pd.DataFrame(Data_Tkt_count_And_Proirities[5])
    Avaya = pd.DataFrame(Avaya_Tkt_count_And_Proirities[6])

    Breach_board =pd.concat([BCG,Avaya,On_Board], ignore_index=True)
    with open("Shift_breaches.html", 'w') as my_file:
        my_file.write('''<html>

        <head>
        	<!-- Latest compiled and minified CSS -->
        <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" integrity="sha384-BVYiiSIFeK1dGmJRAkycuHAHRg32OmUcww7on3RYdg4Va+PmSTsz/K68vbdEjh4u" crossorigin="anonymous">
        <!-- Optional theme -->
        <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap-theme.min.css" integrity="sha384-rHyoN1iRsVXV4nD0JutlnGaslCJuC7uwjduW9SVrLvRYooPp2bWYgmgJQIXwl/Sp" crossorigin="anonymous">

        <title> Handover</title>
        </head>


        <body>
        	<div class="conatiner">
        		<center>
        		<h1>EVENT MANAGEMENT HANDOVER BREACHES</h1>

        		<br>	
        		<br>
        		<div class="table-responsive col-md-12">''')
        my_file.write(Breach_board.to_html(classes='table table-striped'))

        my_file.write('''
        </div></div>
<br><br>

	<!-- Latest compiled and minified JavaScript -->
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js" integrity="sha384-Tc5IQib027qvyjSMfHjOMaLkfuWVxZxUPnCJA7l2mCWNIpG9mGCD8wGNIcPD7Txa" crossorigin="anonymous"></script>
</body>

   </html> 
    ''')
        my_file.close()
    return len(Breach_board)

def send_mail(subject,path,gilbane,allclients,date_now,shift,loginusername):
    import smtplib
    import os
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart
    from email.mime.application import MIMEApplication
    username ='venkatkishore610@outlook.com'
    password ='Kishore@1611'
    fromaddr = username
    toaddr = "venkat.kishore610@gmail.com"


    #dir_path = "C:\\Users\\"+loginusername+"\Desktop\dist\\reports\\"
    dir_path ="E:\Python\python work space\webScraping\\report\\"
    files = "Shift_handover.html"
    img_dir = path
    img_file = gilbane
    img_file2 =allclients

    msg = MIMEMultipart()
    msg['To'] = toaddr
    msg['From'] = fromaddr
    msg['Subject'] = subject

    body = MIMEText('Hello Team, Please find the Event Management Handover report of '+date_now +" "+shift+' shift\n\n\n\n'
                                                                                                           'Thanks&Regards,\n\n'+loginusername)
    msg.attach(body)  # add message body (text or html)

      # add files to the message
    file_path = os.path.join(dir_path, files)
    attachment = MIMEApplication(open(file_path, "rb").read(), _subtype="txt")
    attachment.add_header('Content-Disposition', 'attachment', filename=files)
    msg.attach(attachment)


    file_path = os.path.join(img_dir, gilbane)
    attachment = MIMEApplication(open(file_path, "rb").read(), _subtype="txt")
    attachment.add_header('Content-Disposition', 'attachment', filename=gilbane)
    msg.attach(attachment)

    file_path = os.path.join(img_dir, allclients)
    attachment = MIMEApplication(open(file_path, "rb").read(), _subtype="txt")
    attachment.add_header('Content-Disposition', 'attachment', filename=allclients)
    msg.attach(attachment)


    smtp= smtplib.SMTP('smtp.live.com',587)
    smtp.ehlo()
    smtp.starttls()
    smtp.ehlo()
    smtp.login(username, password)
    smtp.sendmail(msg['From'], msg['To'], msg.as_string())
    smtp.close()

Shift_members = Shift_names(Ticket_Report_Name,reports_path)
Employee = names(Shift_members)
shift = shifts(shift_time)
Data_ALert_count = Data_Alert_Report(Data_Alert_Name,shift,date_now,reports_path)
Avaya_Alert_count = Avaya_Alert_Report(Avaya_Alert_Name,shift,date_now,reports_path)
Avaya_Tkt_count_And_Proirities = Avaya_Ticket_Report(Ticket_Report_Name,Employee[0],shift,date_now,reports_path)
Data_Tkt_count_And_Proirities = Data_Ticket_Report(Ticket_Report_Name,Employee[1],shift,date_now,reports_path)
Total_ticket_count =sla_dashboard(Avaya_Tkt_count_And_Proirities, Data_Tkt_count_And_Proirities)
Priority_Grand_Total =grand_total(Avaya_Tkt_count_And_Proirities, Data_Tkt_count_And_Proirities)
grand_total_sla =grand_total_sla(Avaya_Tkt_count_And_Proirities, Data_Tkt_count_And_Proirities)
Sla_Percentages= sla_dashboard_percentage(Avaya_Tkt_count_And_Proirities, Data_Tkt_count_And_Proirities,Total_ticket_count)


Present_shift_Mems = Employee[0] + Employee[1]
Next_Shift_Mems= Next_Shift_Names + " / "+Next_Esc_manager


Breaches_count =Breaches(Avaya_Tkt_count_And_Proirities,Data_Tkt_count_And_Proirities)
if Breaches_count>0:
    print("There are few incidents breached")
    driver = webdriver.Chrome("E:\Tester\selenium\chromedriver.exe")
    driver.maximize_window()
    time.sleep(3)
    file = "E:\Python\python work space\webScraping\\report\Shift_breaches.html"
    driver.get(file)
    time.sleep(4)
    driver.close()
    driver.quit()
    option =input("Proceed with report..? ")
    if option == "Yes" or "yes":
        print("Inserting Data to Template..")
        template(date_now, shift, Present_shift_Mems,ShiftEm,Next_Shift_Mems, Avaya_Alert_count, Data_ALert_count,
                 Avaya_Tkt_count_And_Proirities, Data_Tkt_count_And_Proirities, Total_ticket_count,
                 Priority_Grand_Total,
                 grand_total_sla,Sla_Percentages)
    elif option == "NO" or "no":
        print("Okay no report created try again")


elif Breaches_count == 0:
    print("No tickets were Breached rendering output template")
    template(date_now,shift,Present_shift_Mems,ShiftEm,Next_Shift_Mems,Avaya_Alert_count,Data_ALert_count,
             Avaya_Tkt_count_And_Proirities,Data_Tkt_count_And_Proirities,Total_ticket_count,Priority_Grand_Total,
             grand_total_sla,Sla_Percentages)
file = "E:\Python\python work space\webScraping\\report\\Shift_handover.html"
driver = webdriver.Chrome("E:\Tester\selenium\chromedriver.exe")
driver.maximize_window()
time.sleep(3)
driver.get(file)
time.sleep(5)
driver.close()
driver.quit()


Subject = "Event Management Handover " + shift + " Shift of " +" "+date_now
try:
    print("Sending Email..")
    send_mail(Subject,reports_path,gilbane,allclients,date_now,shift,loginusername)
    print("Email Sent.")
except Exception as e:
    print('email failed ' +str(e))