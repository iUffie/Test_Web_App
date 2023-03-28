import openpyxl
from openpyxl.styles import colors
from openpyxl.styles import Font, Color, Alignment, PatternFill
from django.shortcuts import render


# Create your views here.
def index(request):
    if request.POST:
        ##                                Загружаем файл и создаем новый                              ##
        file = request.FILES['input_file']
        book2 = openpyxl.Workbook()
        sheet2 = book2.active
        book = openpyxl.load_workbook(file)
        sheet = book.active
        ##                                Создаем шапку                              ##

        ##                  Стили ячеек                   ##
        sheet2.cell(row=1, column=1).value = 'Филиал'
        sheet2.cell(row=1, column=2).value = 'Сотрудник'
        sheet2.cell(row=1, column=3).value = 'Налоговая база'
        sheet2.cell(row=1, column=4).value = 'Налог'
        sheet2.cell(row=2, column=4).value = 'Исчислено всего'
        sheet2.cell(row=2, column=5).value = 'Исчислено всего по формуле'
        sheet2.cell(row=1, column=6).value = 'Отклонения'
        sheet2.merge_cells('A1:A2')
        sheet2.merge_cells('B1:B2')
        sheet2.merge_cells('C1:C2')
        sheet2.merge_cells('F1:G2')
        sheet2.merge_cells('D1:E1')

        for j in range(2):
            for i in range(7):
                sheet2.cell(row=j + 1, column=i + 1).font = Font(bold=True, color='010775', size=10, name='Arial')
                sheet2.cell(row=j + 1, column=i + 1).fill = PatternFill('solid', fgColor="c9e6e4")
                sheet2.cell(row=j + 1, column=i + 1).alignment = Alignment(horizontal='center', vertical="center")
        ##                  Ширина колонок                   ##
        sheet2.column_dimensions['A'].width = 40
        sheet2.column_dimensions['B'].width = 37
        sheet2.column_dimensions['C'].width = 13
        sheet2.column_dimensions['D'].width = 13
        sheet2.column_dimensions['E'].width = 17
        ##                                Парсим эксель и собираем новый  лист                       ##
        for row in range(3, sheet.max_row):
            sheet2.cell(row=row, column=1).value = sheet.cell(row=row, column=1).value
            sheet2.cell(row=row, column=2).value = sheet.cell(row=row, column=2).value
            sheet2.cell(row=row, column=3).value = sheet.cell(row=row, column=5).value
            sheet2.cell(row=row, column=4).value = sheet.cell(row=row, column=6).value
            try:
                if float(sheet.cell(row=row, column=5).value) < 5000000:
                    sheet2.cell(row=row, column=5).value = sheet.cell(row=row, column=5).value * 0.13
                else:
                    sheet2.cell(row=row, column=5).value = sheet.cell(row=row, column=5).value * 0.15
                sheet2.cell(row=row, column=6).value = sheet.cell(row=row, column=6).value - sheet2.cell(row=row,
                                                                                                         column=5).value
            except Exception as ex:
                print(ex)
                sheet2.cell(row=row, column=5).value = 'NaN'
                sheet2.cell(row=row, column=5).value = 'NaN'


    return render(request, 'html_form/index.html')
