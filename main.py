import pandas as pd
import os
import glob


def back(start_input=2):
    start= start_input - 1
    
    # use glob to get all the csv files
    # in the folder
    path = os.getcwd()
    csv_files = glob.glob(os.path.join(path, "*.xlsx"))
    
    each_sum = []
    file_names = []
    
    # loop over the list of csv files
    for f in csv_files:
        
        # read the csv file

        
        sheets_dict = pd.read_excel(f
                           ,sheet_name = None
                           )
        df = pd.concat(sheets_dict)
        #print(df)
        #print(type(df))
        # print the location and filename
        #print('Location:', f)
        
        #print('File Name:', f.split("\\")[-1])
        file_names.append(f.split("\\")[-1])
        # finding sum over index axis
        # By default the axis is set to 0
        each_sum.append( df.sum(axis = 0, skipna = True))
        
        # print the content
        #print('Content:')
        #print(df)
        #print()
    column_names = df.columns
    
    Employe_name = []
    for i in range(start,len(column_names)):
        Employe_name.append(column_names[i])

    incomes = [0] * (len(each_sum[0])-start)
    
    for each in each_sum:
        for i in range(start,len(each)):
                incomes[i-start] += each[i]
    
    income_dict = dict(zip(Employe_name, incomes))
    income_list = list(income_dict.items())
    
    return income_list , file_names, each_sum



while True:
    
    print("EasyExcelSum By:@Pouya.exe")
    print("------------------")
    s = int(input("Enter the column number: "))
    x , y, z= back(s)

    print("Calculation is done!")
    print(x)
    print("File names :")
    print(y)

    s = input("Do you want to see the each sheet summary?[Y/N]: ")
    if s == "Y" or s == "y":    
        print("Each sum :")
        print(z)    
    
    s = input("Colse the app?[Y/N]: ")
    if s == 'Y' or s == 'y':
        break
    else:
        continue
    

