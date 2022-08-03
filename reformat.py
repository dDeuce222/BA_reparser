import xlsxwriter
import math
import pandas as pd

def putdata(sx,ex,sy,ey):
    i = sx 
    j = sy
    try:
        time = data[sx-1][ey]
    except:
        return
    while(i < ex):
        print(sx,ex,sy,ey)
        j = sy
        while(j <= ey):
            selection = float(data[i][j])
            value = float(data[i+1][j])
            if(not math.isnan(selection)  and not math.isnan(value)):
                result_data.append({'Time' : time,'Selection' : selection,'Value' : value,'Target' : value + 1})
            j += 1
        
        i += 2


def reformat(x,y):
    
    if(y >= len(data[0])):
        return
    elif(x >= len(data)):
        # putdata(x,x+3,y,y+2)
        # putdata(x,x+3,y+5,y+7)
        reformat(6,y+10)
    else:
        putdata(x,x+3,y,y+2)
        putdata(x,x+3,y+5,y+7)
        reformat(x+6,y)
    
    # for i in range(stax,endx):
    #     for j in range(stay,endy):
    #         print(data[i][j])

def body():

    loc = input("Please insert refromatted excel file path(ex : parsing sheet.xlsx) : ")
    #sheet = input("Please insert sheetname : ")
    sheet = 'Sheet4'
    #read xlsx file into panda dataframe 
    df = pd.read_excel(loc,engine="openpyxl",sheet_name=sheet)
    #get column names
    global data,result_data
    result_data = []
    data = df.to_numpy()
    i= 6
    j= 0
    reformat(i,j)



    result_df = pd.DataFrame(result_data)
    save_name = input('Please insert save file name(ex : reformateed.xlsx) : ')
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(save_name, engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
    result_df.to_excel(writer, sheet_name='Sheet1', index=False)

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()
    return save_name

body()