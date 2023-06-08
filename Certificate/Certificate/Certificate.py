import csv
def gender_by_patronymic(patronymic):
    male_suffixes = ['ович', 'евич', 'ич']
    if any(patronymic.endswith(suffix) for suffix in male_suffixes):
        return 'м'
    else:
        return 'ж'

def adv(row):
    if gender_by_patronymic(row[3]) == 'м':
        gender_end = "Принял"
    else:
        gender_end = "Приняла"
    return gender_end

with open("Расположение базы данных.csv", newline='', encoding='utf-8') as csvfile:
    reader = csv.reader(csvfile)
    with open("Расположение шаблона сертификата.tex", 'r', encoding='utf-8') as input_file:
        input_str = input_file.read()
        i=0
        for row in reader:
            output_str=input_str.replace('Surname',row[1]).replace('Name',row[2]).replace('Patronim',row[3])\
            .replace('Participated',adv(row))
            with open(f'certificate{i}.tex', 'w') as output_file:
                output_file.write(output_str)
            i+=1

