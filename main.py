import os
os.environ["KMP_DUPLICATE_LIB_OK"]="TRUE"
import xlrd
from time import *
import pandas as pd
import openpyxl

file_dir = "./dataALLFields"
all_file_list=os.listdir(file_dir)

file_injection = []
file_production = []
file_extra = []

path = "./dataAllFields/"

for single_file in all_file_list:
    if single_file.find("Production") >= 0:
        file_production.append(single_file)
    elif single_file.find("Injection") >= 0:
        file_injection.append(single_file)
    else:
        file_extra.append(single_file)

file_production_new = []
file_injection_new = []

i=0
begin_time_p = time()
for single_file in file_production:
    test = xlrd.open_workbook(path+single_file)
    sheet_object = test.sheet_by_name("Well Production")
    nrows = sheet_object.nrows
    i = i+1
    if nrows > 4:
        print(i)
        file_production_new.append(single_file)
    else:
        continue
end_time_p = time()

print("---------------------------------")

print("The length of file_production_new is %d" % len(file_production_new))
print("The running time of file_production_new is %f" % (end_time_p - begin_time_p))


print("---------------------------------")

res_p = pd.DataFrame(columns=['Production Date', 'Oil Produced (bbl)', 'Water Produced (bbl)', 'Well #', 'Latitude', 'Longitude'])

begin_time_p_new = time()
p = 0
for single_file in file_production_new:
    data1 = pd.read_excel(path+single_file, header=3, usecols=[1, 2,3])
    data2 = pd.read_excel(path+single_file, header=0, usecols=[5, 13, 14])
    data1["Well #"] = data2.iat[0, 0]
    data1["Latitude"] = data2.iat[0, 1]
    data1["Longitude"] = data2.iat[0, 2]
    res_p = pd.concat([res_p, data1], axis=0, ignore_index=True)
    p = p+1
    print(p)
end_time_p_new = time()

time_p = end_time_p_new - begin_time_p_new
print("---------------------------------")
print(time_p)
print()
print(res_p)
