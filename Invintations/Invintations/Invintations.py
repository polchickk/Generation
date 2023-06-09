import csv

def positionData(row):
    work = row[6].split()
    for i in range(len(work)):
        if work[i][-2] == 'ы' and work[i][-1] == 'й':
            work[i] = work[i][:-2] + 'ому'
        elif work[i][-1] == 'т' or work[i][-1] == 'к':
            work[i] = work[i] + 'у'
    position = ' '.join(work)
    return position

def affiliationGen(row):
    job = row[5].split()
    for i in range(len(job)):
        if len(job[i]) >= 2:
            if job[i][-2] == 'ы' and job[i][-1] == 'й':
                job[i] = job[i][:-2] + 'ого'
            elif job[i][-2] == 'и' and job[i][-1] == 'й':
                job[i] = job[i][:-2] + 'ого'
            elif job[i][-1] == 'т':
                job[i] = job[i] + 'а'
    institute = ' '.join(job)
    return institute

def surnameData(row):
    surname = row[1].split()[-1]
    if surname.endswith('ов') or surname.endswith('ев') or surname.endswith('ин') or surname.endswith('ын'):
        surname += 'у'
    elif surname.endswith('ина') or surname.endswith('ына') or surname.endswith('ова') or surname.endswith(
            'ева'):
        surname = surname[:-1] + 'ой'
    elif surname.endswith('ий') or surname.endswith('ый') or surname.endswith('ей') or surname.endswith('ой'):
        surname = surname[:-2] + 'ому'
    elif surname.endswith('ая'):
        surname = surname[:-2] + 'ой'
    elif surname.endswith('яя'):
        surname = surname[:-2] + 'ей'
    else:
        surname = surname
    return surname


def gender_by_patronymic(patronymic):
    male_suffixes = ['ович', 'евич', 'ич']
    if any(patronymic.endswith(suffix) for suffix in male_suffixes):
        return 'м'
    else:
        return 'ж'

def adv(row):
    if gender_by_patronymic(row[3]) == 'м':
        gender_end = "Уважаемый"
    else:
        gender_end = "Уважаемая"
    return gender_end

with open("Расположение базы данных.csv", newline='', encoding='utf-8') as csvfile:
    reader = csv.reader(csvfile)
    with open("Расположение шаблона приглашения.tex", 'r', encoding='utf-8') as input_file:
        input_str = input_file.read()
        i=0
        for row in reader:
            pos = input_str.replace('PositionDat', positionData(row)).replace('AffiliationGen', affiliationGen(row))\
            .replace('SurnameDat',surnameData(row)).replace('NameF', row[2][0]).replace('PatronimF', row[3][0]).replace('Уважаемый',adv(row))\
            .replace('Name',row[2]).replace('Patronim',row[3])

            with open(f'invintation{i}.tex', 'w') as output_file:
                output_file.write(pos)
            i+=1

