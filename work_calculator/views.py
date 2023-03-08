# Create your views here.
from django.contrib.auth.decorators import login_required
from django.http import HttpResponse
from django.shortcuts import render
import openpyxl
from work_calculator.src.work_calculator import WorkCalculator
import xlwt

@login_required
def index(request):
    if "GET" == request.method:
        return render(request, 'work_calculator/index.html', {})
    else:
        excel_file = request.FILES["excel_file"]
        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="file.xlsx"'
        # you may put validations here to check extension or file size

        wb = openpyxl.load_workbook(excel_file)
        wb.save('file.xlsx')
        wc = WorkCalculator('file.xlsx', 'output.xlsx')
        wc.create_report()
        # getting a particular sheet by name out of many sheets

        wb = openpyxl.load_workbook('output.xlsx')
        wb.save(response)
        return response
