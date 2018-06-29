# -*- coding: utf-8 -*-

from openpyxl import load_workbook
from collections import OrderedDict
import operator
import csv

from models.patent_record import Patent
from utils import extractMedicines, extractFormOfPatent, extractFunctionFromPatent, extractMajorFunctionFromPatent

# 1. load excel file
wb = load_workbook("dataset/Copy of 桔梗专利汇总-6.28-2桔梗已替换.xlsx")

# 2. get all sheets of this excel file
print("sheet names:", wb.sheetnames)

# 3. Sheet1

sheet1 = wb['Sheet1']
sheet2 = wb['Sheet2']

"""
    Process data of Sheet1
"""

total_medicines = []
total_forms = []
total_major_functions = []
total_functions = []

total_medicines_set = set()
total_forms_set = set()
total_major_functions_set = set()
total_functions_set = set()

patent_medicines = []

for i in range(2, 5964):
    print(i)

    # if i == 10:
    #     break

    # patent name
    patent_name = ''
    if sheet1['A' + str(i)].value:
        patent_name = sheet1['A' + str(i)].value.strip()

    # medicine components
    medicine_components = ''
    medicines = []
    if sheet1['B'+str(i)].value:
        medicine_components = sheet1['B' + str(i)].value.strip()
        medicines = extractMedicines(medicine_components)
        # print(medicines)
        if medicines is not None and len(medicines) > 0:
            patent_medicines.append(medicines)



    # patent form
    form_str = ''
    forms = []
    if sheet1['C'+str(i)].value:
        form_str = sheet1['C' + str(i)].value.strip()
        forms = extractFormOfPatent(form_str)
        # print(forms)

    # patent major function
    major_function_str = ''
    major_functions = []
    if sheet1['D'+str(i)].value:
        major_function_str = sheet1['D' + str(i)].value.strip()
        major_functions = extractMajorFunctionFromPatent(major_function_str)
        # print(major_functions)

    # patent functions
    function_str = ''
    functions = []
    if sheet1['E'+str(i)].value:
        function_str = sheet1['E' + str(i)].value.strip()
        functions = extractFunctionFromPatent(function_str)
        # print(functions)

    # patent id
    patent_id = ''
    if sheet1['F'+str(i)].value:
        patent_id = str(sheet1['F'+str(i)].value).strip()

    # total data
    total_medicines += medicines
    total_forms += forms
    total_major_functions += major_functions
    total_functions += functions

print("total medicines num: %d" % len(total_medicines))
print("total forms num: %d" % len(total_forms))
print("total major functions num: %d" % len(total_major_functions))
print("total functions num: %d" % len(total_functions))

total_medicines_set = set(total_medicines)
total_forms_set = set(total_forms)
total_major_functions_set = set(total_major_functions)
total_functions_set = set(total_functions)
print("total medicines set num: %d" % len(total_medicines_set))
print("total forms set num: %d" % len(total_forms_set))
print("total major functions set num: %d" % len(total_major_functions_set))
print("total functions num: %d" % len(total_functions_set))

# statistics the medicines
total_medicines_map = {}

for med in total_medicines_set:
    if total_medicines.count(med) > 0:
        total_medicines_map[med] = total_medicines.count(med)

# sort map
total_medicines_sorted = sorted(total_medicines_map.items(), key=operator.itemgetter(1), reverse=True)
print(total_medicines_sorted)

# statistics the forms
total_forms_map = {}
for form in total_forms_set:
    if total_forms.count(form) > 0:
        total_forms_map[form] = total_forms.count(form)
# sort forms map
total_forms_sorted = sorted(total_forms_map.items(), key=operator.itemgetter(1), reverse=True)
print(total_forms_sorted)

# statistics the major functions
total_major_functions_map = {}
for mj in total_major_functions_set:
    if total_major_functions.count(mj) > 0:
        total_major_functions_map[mj] = total_major_functions.count(mj)
# sort major function map
total_major_functions_sorted = sorted(total_major_functions_map.items(), key=operator.itemgetter(1), reverse=True)
print(total_major_functions_sorted)


# statistics the functions
total_functions_map = {}
for func in total_functions_set:
    if total_functions.count(func) > 0:
        total_functions_map[func] = total_functions.count(func)
# sorted functions map
total_functions_sorted = sorted(total_functions_map.items(), key=operator.itemgetter(1), reverse=True)
print(total_functions_sorted)

# write to .csv file
# 1. total medicines sorted
with open("中药统计.csv", "w", encoding='utf-8-sig') as f:
    writer = csv.writer(f)
    writer.writerow(("中药名称", "数量"))
    for item in total_medicines_sorted:
        writer.writerow(item)

# 2. total forms sorted
with open("剂型统计.csv", "w", encoding="utf-8-sig") as f:
    writer = csv.writer(f)
    writer.writerow(("剂型", "数量"))
    for item in total_forms_sorted:
        writer.writerow(item)


# 3. total major functions sorted
with open("主治统计.csv", "w", encoding="utf-8-sig") as f:
    writer = csv.writer(f)
    writer.writerow(("主治", "数量"))
    for item in total_major_functions_sorted:
        writer.writerow(item)

# 4. total functions sorted
with open("功能统计.csv", "w", encoding="utf-8-sig") as f:
    writer = csv.writer(f)
    writer.writerow(("功能", "数量"))
    for item in total_functions_sorted:
        writer.writerow(item)

with open("patent_medicines.csv", "w", encoding="utf-8-sig") as f:
    for record in patent_medicines:
        for i in range(len(record)):
            if i == len(record) - 1:
                f.write(record[i])
            else:
                f.write(record[i] + ",")
        f.write("\n")

print("Sheet 1 data processed sucessed!")

"""
    Process data of Sheet2
"""

total_medicines = []
total_forms = []
total_major_functions = []
total_functions = []

total_medicines_set = set()
total_forms_set = set()
total_major_functions_set = set()
total_functions_set = set()

patent_medicines = []

for i in range(2, 471):
    print(i)

    # if i == 10:
    #     break

    # patent name
    patent_name = ''
    if sheet2['A' + str(i)].value:
        patent_name = sheet2['A' + str(i)].value.strip()

    # medicine components
    medicine_components = ''
    medicines = []
    if sheet2['B'+str(i)].value:
        medicine_components = sheet2['B' + str(i)].value.strip()
        medicines = extractMedicines(medicine_components)
        # print(medicines)
        if medicines is not None and len(medicines) > 0:
            patent_medicines.append(medicines)



    # patent form
    form_str = ''
    forms = []
    if sheet2['C'+str(i)].value:
        form_str = sheet2['C' + str(i)].value.strip()
        forms = extractFormOfPatent(form_str)
        # print(forms)

    # patent major function
    major_function_str = ''
    major_functions = []
    if sheet2['D'+str(i)].value:
        major_function_str = sheet2['D' + str(i)].value.strip()
        major_functions = extractMajorFunctionFromPatent(major_function_str)
        # print(major_functions)

    # patent functions
    function_str = ''
    functions = []
    if sheet2['E'+str(i)].value:
        function_str = sheet2['E' + str(i)].value.strip()
        functions = extractFunctionFromPatent(function_str)
        # print(functions)

    # patent id
    patent_id = ''
    if sheet2['F'+str(i)].value:
        patent_id = str(sheet2['F'+str(i)].value).strip()

    # total data
    total_medicines += medicines
    total_forms += forms
    total_major_functions += major_functions
    total_functions += functions

print("total medicines num: %d" % len(total_medicines))
print("total forms num: %d" % len(total_forms))
print("total major functions num: %d" % len(total_major_functions))
print("total functions num: %d" % len(total_functions))

total_medicines_set = set(total_medicines)
total_forms_set = set(total_forms)
total_major_functions_set = set(total_major_functions)
total_functions_set = set(total_functions)
print("total medicines set num: %d" % len(total_medicines_set))
print("total forms set num: %d" % len(total_forms_set))
print("total major functions set num: %d" % len(total_major_functions_set))
print("total functions num: %d" % len(total_functions_set))


# statistics the medicines
total_medicines_map = {}

for med in total_medicines_set:
    if total_medicines.count(med) > 0:
        total_medicines_map[med] = total_medicines.count(med)

# sort map
total_medicines_sorted = sorted(total_medicines_map.items(), key=operator.itemgetter(1), reverse=True)
print(total_medicines_sorted)

# statistics the forms
total_forms_map = {}
for form in total_forms_set:
    if total_forms.count(form) > 0:
        total_forms_map[form] = total_forms.count(form)
# sort forms map
total_forms_sorted = sorted(total_forms_map.items(), key=operator.itemgetter(1), reverse=True)
print(total_forms_sorted)

# statistics the major functions
total_major_functions_map = {}
for mj in total_major_functions_set:
    if total_major_functions.count(mj) > 0:
        total_major_functions_map[mj] = total_major_functions.count(mj)
# sort major function map
total_major_functions_sorted = sorted(total_major_functions_map.items(), key=operator.itemgetter(1), reverse=True)
print(total_major_functions_sorted)


# statistics the functions
total_functions_map = {}
for func in total_functions_set:
    if total_functions.count(func) > 0:
        total_functions_map[func] = total_functions.count(func)
# sorted functions map
total_functions_sorted = sorted(total_functions_map.items(), key=operator.itemgetter(1), reverse=True)
print(total_functions_sorted)

# write to .csv file
# 1. total medicines sorted
with open("茶剂-中药统计.csv", "w", encoding='utf-8-sig') as f:
    writer = csv.writer(f)
    writer.writerow(("中药名称", "数量"))
    for item in total_medicines_sorted:
        writer.writerow(item)

# 2. total forms sorted
with open("茶剂-剂型统计.csv", "w", encoding="utf-8-sig") as f:
    writer = csv.writer(f)
    writer.writerow(("剂型", "数量"))
    for item in total_forms_sorted:
        writer.writerow(item)


# 3. total major functions sorted
with open("茶剂-主治统计.csv", "w", encoding="utf-8-sig") as f:
    writer = csv.writer(f)
    writer.writerow(("主治", "数量"))
    for item in total_major_functions_sorted:
        writer.writerow(item)

# 4. total functions sorted
with open("茶剂-功能统计.csv", "w", encoding="utf-8-sig") as f:
    writer = csv.writer(f)
    writer.writerow(("功能", "数量"))
    for item in total_functions_sorted:
        writer.writerow(item)

with open("tea_patent_medicines.csv", "w", encoding="utf-8-sig") as f:
    for record in patent_medicines:
        for i in range(len(record)):
            if i == len(record) - 1:
                f.write(record[i])
            else:
                f.write(record[i] + ",")
        f.write("\n")

print("Sheet 1 data processed sucessed!")
