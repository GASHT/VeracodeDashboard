__author__ = 'Gaurav Sharma'
import pandas as pd
import xlsxwriter
import requests
import xml.etree.ElementTree as ET

username = "bermuda_api"
password = "0mIz25Knfg7IDsft5p4K7sSTEfOTmj"

def main():

    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook('Veracode.xlsx')
    worksheet = workbook.add_worksheet()

    worksheet.write('A1', 'App_ID')
    worksheet.write('B1', 'Build_Id')
    worksheet.write('C1', 'App_Name')
    worksheet.write('D1', 'Submitted_Date')
    worksheet.write('E1', 'Published_Date')
    worksheet.write('F1', 'Policy_Compliance_Status')
    worksheet.write('G1', 'Rating')
    worksheet.write('H1', 'Score')

    # Add a bold format to use to highlight cells.
    cell_format = workbook.add_format({'bold': True, 'font_color': 'blue'})
    cell_format.set_align('center')

    # Add a number format for cells with Score.
    Score = workbook.add_format({'num_format': '$#,##0'})
    App_ID = workbook.add_format({'num_format': '$#,##0'})
    Build_Id = workbook.add_format({'num_format': '$#,##0'})

    # Add a cell format for cells.
    worksheet.set_row(0, 20)
    worksheet.set_column('A:A', 15)
    worksheet.set_column('B:B', 15)
    worksheet.set_column('C:C', 60)
    worksheet.set_column('D:D', 25)
    worksheet.set_column('E:E', 25)
    worksheet.set_column('F:F', 25)
    worksheet.set_column('G:G', 15)
    worksheet.set_column('H:H', 15)

    # Add a cell alignment for cells.

    # Write some data headers.
    worksheet.write('A1', 'App_ID', cell_format)
    worksheet.write('B1', 'Build_Id', cell_format)
    worksheet.write('C1', 'App_Name', cell_format)
    worksheet.write('D1', 'Submitted_Date', cell_format)
    worksheet.write('E1', 'Published_Date', cell_format)
    worksheet.write('F1', 'Policy_Compliance_Status', cell_format)
    worksheet.write('G1', 'Rating', cell_format)
    worksheet.write('H1', 'Score', cell_format)

    # Start from the first cell below the headers.
    row = 1
    col = 0

    print("Getting List of App_ID & Their Corresponding Build_ID:")

    xmlApplist = getApplist()
    applist = ET.fromstring(xmlApplist)
    try:
     for appnode in applist:
        print("app_id")
        print(appnode.get('app_id'))
        xmlBuildlist = getBuildlist(appnode.get('app_id'))
        root = ET.fromstring(xmlBuildlist)
        print("Getting Last build_id of Application:")
        print(root[0].get('build_id'))
        xmlBuildinfo = getBuildInfo(appnode.get('app_id'))
        root = ET.fromstring(xmlBuildinfo)
        app_id = ""
        build_id = ""
        app_name = ""
        submitted_date = ""
        published_date = ""
        policy_compliance_status = ""
        rating = ""
        score = ""
        if (root[0].get('results_ready') == 'false'):
            print("extracting build results below:")
            print('result Ready:' + root[0].get('results_ready') + ', App_id : ' + root.get('app_id') + ', Build_id : ' + root.get('build_id') + ',status : ' + root[0][0].get("status"))
            worksheet.write_string(row, col, root.get('app_id'))
            worksheet.write_string(row, col + 1, root.get('build_id'))
            worksheet.write_string(row, col + 2, appnode.get('app_name'))
            worksheet.write_string(row, col + 3, "NA")
            worksheet.write_string(row, col + 4, "NA")
            worksheet.write_string(row, col + 5, root[0][0].get('status'))
            worksheet.write_string(row, col + 6, "NA")
            worksheet.write_string(row, col + 7, "NA")
            row += 1
        else:
             print("getting build details")
             xmlBuildDetail = getXMLSummaryReport(root[-1].get('build_id'))
             root = ET.fromstring(xmlBuildDetail)
             print('App_id : ' + root.get('app_id') + ', Build_id : ' + root.get('build_id') + ', App_name : ' + root.get('app_name') + ', Submitted_Date : ' + root[0].get('submitted_date') + ', Published_Date : ' + root[0].get('published_date') + ', Policy_compliance_status : ' + root.get('policy_compliance_status') + ", Rating: " + root[0].get('rating') + ", Score: " + root[0].get('score'))
             worksheet.write_string(row, col, root.get('app_id'))
             worksheet.write_string(row, col + 1, root.get('build_id'))
             worksheet.write_string(row, col + 2, root.get('app_name'))
             worksheet.write_string(row, col + 3, root[0].get('submitted_date'))
             worksheet.write_string(row, col + 4, root[0].get('published_date'))
             worksheet.write_string(row, col + 5, root.get('policy_compliance_status'))
             worksheet.write_string(row, col + 6, root[0].get('rating'))
             worksheet.write_string(row, col + 7, root[0].get('score'))
             row+=1
     workbook.close()
     exceltohtml()
    except IndexError:
            print('No childNodes for this element')

# Load the excel file and read as xls

def getXMLSummaryReport(build_id):
    PARAMS = {'build_id': build_id}
    r = requests.post("https://analysiscenter.veracode.com/api/2.0/summaryreport.do", data=PARAMS, auth=(username, password))
    return r.content


def getBuildInfo(app_id):
    PARAMS = {'app_id': app_id}
    r = requests.post("https://analysiscenter.veracode.com/api/5.0/getbuildinfo.do", data=PARAMS, auth=(username, password))
    return r.text

def getBuildlist(app_id):
    PARAMS = {'app_id': app_id}
    r = requests.post("https://analysiscenter.veracode.com/api/4.0/getbuildlist.do", data=PARAMS, auth=(username, password))
    return r.text

def getApplist():
    r = requests.post("https://analysiscenter.veracode.com/api/5.0/getapplist.do", auth=(username, password))
    return r.text

def exceltohtml():
    xls_file = pd.ExcelFile('Veracode.xlsx')
    xls_file.sheet_names
    # Load the xls file's Sheet1 as a dataframe
    df = xls_file.parse('Sheet1')
    # Load the dataframe file's Sheet1 as a HTML
    df.to_html('Veracode.html')
    return;

if __name__ == "__main__":
    main()
