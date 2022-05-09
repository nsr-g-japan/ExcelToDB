from django.shortcuts import render
from openpyxl import load_workbook
from django.http import HttpResponseRedirect
import xlrd
import pandas
import pyodbc
import pyodbc
import pandas as pd

db = pyodbc.connect('Driver={SQL server};' 'server=Nisarbasha;' 'Database=Testdata;' 'Trusted_connection=yes;')
cursor = db.cursor()


# Create your views here.
def homepage(request):
    if request.method == "POST":
        global getsheet,sheets, datas
        getsheet = request.POST.get('sheetpath')
        print(getsheet)
        ws = load_workbook('{}'.format(getsheet))
        wb = xlrd.open_workbook(getsheet)
        sheets = ws.get_sheet_names()
        datas = {}
        for sheet in sheets:
            sheet1 = wb.sheet_by_name(sheet)
            row = sheet1.row(0)
            table = []
            for mycell in row:
                table.append(mycell.value)
            datas.update({sheet: table})
        return HttpResponseRedirect('sheetname')
    return render(request, 'homepage.html', {})


def sheetname(request, ):
    if request.method == "POST":
        global datasheet
        datasheet = request.POST.getlist('sheet[]')
        for tbl in datasheet:
            clm = request.POST.getlist("cval['{0}'][]".format(tbl))
            clset = ' '
            for cl in clm:
                dt = request.POST.get("val['{0}']['{1}']".format(tbl, cl))
                clset += cl + ' ' + dt + ', '
            cursor.execute('if OBJECT_ID(\'{0}\') IS NOT NULL DROP TABLE {0}'.format(tbl))
            cursor.execute('create table {}({})'.format(tbl, clset))
            db.commit()

            print("------------")
    return render(request, 'sheetlist.html', {'data': datas})
