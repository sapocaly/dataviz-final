import xlrd
import csv
from os import listdir
from os.path import isfile, join

def Excel2CSV(ExcelFile, SheetName):
     workbook = xlrd.open_workbook(ExcelFile)
     worksheet = workbook.sheet_by_name(SheetName)
     name =  worksheet.cell(1, 0).value.replace('/',',')
     csvfile = open(name + '.csv','wb')
     wr = csv.writer(csvfile, quoting=csv.QUOTE_ALL)

     for rownum in xrange(worksheet.nrows):
         wr.writerow(
             list(x.encode('utf-8') if type(x) == type(u'') else x
                  for x in worksheet.row_values(rownum)))

     csvfile.close()



def convert_all_to_csv():
    onlyfiles = [f for f in listdir('.') if isfile(join('.', f))]
    onlyxls = [f for f in onlyfiles if f[-3:] == 'xls']
    for xls in onlyxls:
        Excel2CSV(xls,'Partner')


def get_country_list():
    country_list = []
    onlyfiles = [f for f in listdir('.') if isfile(join('.', f))]
    onlycsv = [f for f in onlyfiles if f[-3:] == 'csv']
    for file in onlycsv:
        with open(file) as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                country_list.append(row['Reporter Name'])
                break
    return country_list


def generate_big_dict():
    onlyfiles = [f for f in listdir('.') if isfile(join('.', f))]
    onlycsv = [f for f in onlyfiles if f[-3:] == 'csv']
    global country_list
    big_dict = {}
    for file in onlycsv:
        with open(file) as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                country = row['Reporter Name']
                break
            country_dict = {}
            for c in country_list:
                country_dict[c] = 0
            for row in reader:
                count = row['Export (US$ Thousand)']
                partner = row['Partner Name']
                if partner in country_list:
                    if count != '':
                        country_dict[partner] += float(count)

        big_dict[country] = country_dict

    return big_dict


print '\n\n\n\n\n\n\n'
SouthAsia = ['Afghanistan','Bangladesh','India','Pakistan','Sri Lanka']
NorthAsia = ['Kazakhstan','Georgia','Mongolia','Russian Federation','Turkey','United Arab Emirates']
WestAsia = ['Azerbaijan','Armenia','Bahrain','Israel','Jordan','Kuwait','Oman','Qatar']
SoutheastAsia = ['Brunei','Indonesia','Malaysia','Maldives','Philippines','Singapore','Thailand','Vietnam']
EastAsia = ['China','Hong Kong, China','Japan','Macao']
NonEu = ['Albania','Andorra','Belarus','Bosnia and Herzegovina','Iceland','Macedonia, FYR','Moldova','Montenegro','Norway','Serbia, FR(Serbia/Montenegro)','Switzerland','Ukraine']
Eu = NonEu + ['Austria','Belgium','Bulgaria','Croatia','Cyprus','Czech Republic','Denmark','Estonia','Finland','France','Germany','Greece','Hungary','Ireland','Italy','Latvia','Lithuania','Malta','Luxembourg','Netherlands','Poland','Portugal','Romania','Slovak Republic','Slovenia','Spain','Sweden','United Kingdom']
Africa = ['Algeria','Benin','Botswana','Burkina Faso','Burundi','Cameroon','Cape Verde',"Cote d'Ivoire",'Egypt, Arab Rep.','Ethiopia(excludes Eritrea)','Gambia, The','Mauritania','Mauritius','Mozambique','Niger','Rwanda','Sao Tome and Principe','Senegal','Tanzania','Zambia','Zimbabwe']
SouthAmerica = ['Argentina','Bolivia','Brazil','Chile','Colombia','Ecuador','Guyana','Paraguay','Peru','South Africa','Suriname','Uruguay']
NorthAmerica = ['Aruba','Bahamas, The','Barbados','Belize','Bermuda','Canada','Dominican Republic','El Salvador','Greenland','Guatemala','Jamaica','Mexico','Montserrat','Nicaragua','Panama','United States']
Oceania = ['Australia','Fiji','French Polynesia','New Caledonia','New Zealand','Palau','Samoa','Solomon Islands','Tonga']










def get_continent(country):
    global SouthAsia, NorthAsia, WestAsia, EastAsia, Europe, Africa, SouthAmerica, NorthAmerica, Oceania
    if country in SouthAsia:
        return 'South Asia'
    if country in NorthAsia:
        return 'North Asia'
    if country in WestAsia:
        return 'West Asia'
    if country in SoutheastAsia:
        return 'Southeast Asia'
    if country in EastAsia:
        return 'East Asia'
    if country in Eu:
        return 'Europe'
    if country in Africa:
        return 'Africa'
    if country in SouthAmerica:
        return 'South America'
    if country in NorthAmerica:
        return 'North America'
    if country in Oceania:
        return 'Oceania'
    raise Exception()

country_list = get_country_list()
big_dict = generate_big_dict()
all_countries = ['East Asia'] + EastAsia + ['South Asia'] + SouthAsia + ['West Asia'] + WestAsia + ['North Asia'] + NorthAsia + ['Southeast Asia'] + SoutheastAsia +['Europe'] + Eu + ['Africa'] + Africa + ['North America'] + NorthAmerica + ['South America'] + SouthAmerica + ['Oceania'] + Oceania


index_dict = {}
for i in range(len(all_countries)):
    index_dict[all_countries[i]] = i


indices = [index_dict['East Asia'], index_dict['South Asia'],index_dict['West Asia'],index_dict['North Asia'],index_dict['Southeast Asia'],index_dict['Europe'],index_dict['Africa'],index_dict['North America'],index_dict['South America'],index_dict['Oceania']]

SouthAsia_dict = {}
for country in country_list:
    SouthAsia_dict[country] = 0;


NorthAsia_dict = {}
for country in country_list:
    NorthAsia_dict[country] = 0;

WestAsia_dict = {}
for country in country_list:
    WestAsia_dict[country] = 0;

EastAsia_dict = {}
for country in country_list:
    EastAsia_dict[country] = 0;

SoutheastAsia_dict = {}
for country in country_list:
    SoutheastAsia_dict[country] = 0;

NonEu_dict = {}
for country in country_list:
    NonEu_dict[country] = 0;

Eu_dict = {}
for country in country_list:
    Eu_dict[country] = 0;

Africa_dict = {}
for country in country_list:
    Africa_dict[country] = 0;

SouthAmerica_dict = {}
for country in country_list:
    SouthAmerica_dict[country] = 0;

NorthAmerica_dict = {}
for country in country_list:
    NorthAmerica_dict[country] = 0;

Oceania_dict = {}
for country in country_list:
    Oceania_dict[country] = 0;

for country in country_list:
    small_dict = big_dict[country]
    if country in SouthAsia:
        for c in country_list:
            SouthAsia_dict[c] += small_dict[c]
        continue
    if country in NorthAsia:
        for c in country_list:
            NorthAsia_dict[c] += small_dict[c]
        continue
    if country in WestAsia:
        for c in country_list:
            WestAsia_dict[c] += small_dict[c]
        continue
    if country in EastAsia:
        for c in country_list:
            EastAsia_dict[c] += small_dict[c]
        continue
    if country in SoutheastAsia:
        for c in country_list:
            SoutheastAsia_dict[c] += small_dict[c]
        continue
    if country in Eu:
        for c in country_list:
            Eu_dict[c] += small_dict[c]
        continue
    if country in Africa:
        for c in country_list:
            Africa_dict[c] += small_dict[c]
        continue
    if country in SouthAmerica:
        for c in country_list:
            SouthAmerica_dict[c] += small_dict[c]
        continue
    if country in NorthAmerica:
        for c in country_list:
            NorthAmerica_dict[c] += small_dict[c]
        continue
    if country in Oceania:
        for c in country_list:
            Oceania_dict[c] += small_dict[c]
        continue
    raise Exception()



big_dict['South Asia'] = SouthAsia_dict
big_dict['North Asia'] = NorthAsia_dict
big_dict['West Asia'] = WestAsia_dict
big_dict['East Asia'] = EastAsia_dict
big_dict['Southeast Asia'] = SoutheastAsia_dict
big_dict['Europe'] = Eu_dict
big_dict['Africa'] = Africa_dict
big_dict['South America'] = SouthAmerica_dict
big_dict['North America'] = NorthAmerica_dict
big_dict['Oceania'] = Oceania_dict


for d in all_countries:
    small_dict = big_dict[d]
    small_dict['South Asia'] = 0
    small_dict['North Asia'] = 0
    small_dict['West Asia'] = 0
    small_dict['Southeast Asia'] = 0
    small_dict['East Asia'] = 0
    small_dict['Europe'] = 0
    small_dict['Africa'] = 0
    small_dict['South America'] = 0
    small_dict['North America'] = 0
    small_dict['Oceania'] = 0
    for c in country_list:
        if c in SouthAsia:
            small_dict['South Asia'] += small_dict[c]
            continue
        if c in NorthAsia:
            small_dict['North Asia'] += small_dict[c]
            continue
        if c in WestAsia:
            small_dict['West Asia'] += small_dict[c]
            continue
        if c in EastAsia:
            small_dict['East Asia'] += small_dict[c]
            continue
        if c in SoutheastAsia:
            small_dict['Southeast Asia'] += small_dict[c]
            continue
        if c in Eu:
            small_dict['Europe'] += small_dict[c]
            continue
        if c in Africa:
            small_dict['Africa'] += small_dict[c]
            continue
        if c in SouthAmerica:
            small_dict['South America'] += small_dict[c]
            continue
        if c in NorthAmerica:
            small_dict['North America'] += small_dict[c]
            continue
        if c in Oceania:
            small_dict['Oceania'] += small_dict[c]
    big_dict[d] = small_dict





final_list = []
for c in all_countries:
    current_list = []
    for d in all_countries:
        current_list.append(big_dict[c][d])
    final_list.append(current_list)

print len(all_countries)
print len(final_list)
print len(final_list[0])
print indices

print '---------test---field---------'


import json

json_str = json.dumps({'names':all_countries, 'regions':indices, 'matrix':{'2014':final_list}})

fin = open('output_final.json','w')
fin.write(json_str)
fin.close()












