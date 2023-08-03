# Author: Dalwin Sundar
# Date: Aug 2, 2023
import re #1
import pandas as pd
import os.path
import sys

pathstr = './'
n = len(sys.argv)
if n == 1:
    filename = 'input.xlsx'
else:
    filename = sys.argv[1]
namelength = len(filename)
last5 = filename[namelength - 5:namelength]
pathstr += filename
if os.path.isfile(pathstr) and last5 == '.xlsx':
    df = pd.DataFrame(pd.read_excel(filename))
    df.loc[-1] = list(df.columns)
    df.index = df.index + 1
    df = df.sort_index()
    temp = list(df.columns)
    df = df.rename(columns={temp[0]: 'agency', temp[1]: 'undivided', temp[2]: 'unknown'})
    # df #2
    all_fields = df['undivided'].tolist() #3
    # test = all_fields[54]
    # test
    def get_phone(s): #4
        phone = re.findall('\(\d{3}\)\s\d{3}-\d{4}', s)
        if len(phone) == 0:
            return 'N/A'
        else:
            return phone[0]
    def get_name(s):
        exp = re.findall('[^,]+', s)
        return exp[0]
    def get_email(s):
        email = re.findall('[^\s]+@[^\s]+\.[^\s]+', s)
        if len(email) == 0:
            return 'N/A'
        else:
            return email[0]
    def get_position(s):
        commas = len(re.findall('[^\s]+@[^\s]+\.[^\s]+', s)) + len(re.findall('\(\d{3}\)\s\d{3}-\d{4}', s))
        first = len(get_name(s)) + 1
        while s[first] == " ":
            first += 1
        last = len(s) - 1
        while commas > 0:
            if s[last] == ",":
                commas -= 1
            last -= 1
        last += 1
        return s[first:last]
    # print(get_phone(test)) #5
    # print(get_name(test))
    # print(get_email(test))
    # print(get_position(test))
    name = [] #6
    position = []
    phone = []
    email = []
    for string in all_fields:
        name.append(get_name(string))
        phone.append(get_phone(string))
        email.append(get_email(string))
        position.append(get_position(string))
    df['name'] = name #7
    df['position'] = position
    df['phone'] = phone
    df['email'] = email
    # df
    df = df.drop('undivided', axis=1)
    if filename == 'input.xlsx':
        outfile = 'output.csv'
    else:
        outfile = filename[:namelength - 5] + '.csv'
    df.to_csv(outfile) #8
elif last5 != '.xlsx':
    print('Error: input is not an .xlsx file')
else:
    print('Error: the file ' + filename + ' does not exist in current directory')




