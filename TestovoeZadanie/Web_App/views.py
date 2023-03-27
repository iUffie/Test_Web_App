from django.shortcuts import render
import pandas


# Create your views here.
def index(request):
    if request.POST:

        print(request.POST)
        print(request.FILES)

        file = request.FILES['input_file']
        print(f"file {file}")
        df = pandas.read_excel(file, header=None)
        name_to_pars = ['Филиал', 'Сотрудник', 'Налоговая база', 'Исчислено всего']
        try:
            for el in name_to_pars:
                for col in df.columns:
                    i = 0
                    while i < 2:
                        if el == df[col][i]:
                            print(f"Найдено совпадение {el} {df[col][i]}")
                        i = i + 1
        except Exception as ex:
            print(ex)
    return render(request, 'html_form/index.html')
