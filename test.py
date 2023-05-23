import pandas as pd
from pandas_datareader import data as wb
import openpyxl
from breeze_connect import BreezeConnect

session_token = '1692839'
api_key = '6588c73p34F6^7A91VD4_3c328563!36'
api_secret = '1!3#3OLWo0R^8217y381vq75R015G806'



breeze = BreezeConnect(api_key=api_key)
breeze.generate_session(api_secret=api_secret,session_token=session_token)

strikes = pd.ExcelWriter('Strike.xlsx', engine='xlsxwriter')
option = pd.ExcelWriter('CMP.xlsx', engine='xlsxwriter')

stock_code = "CNXBAN"
interval_1 = "5minute"
interval_2 = "1minute"
from_date = "2021-03-12T07:00:00.000Z"
to_date = "2021-03-18T07:00:00.000Z"
expiry_date="2021-03-18T07:00:00.000Z"

sdata = breeze.get_historical_data(interval=interval_1,
                                   from_date= from_date,
                                   to_date= to_date,
                                   stock_code=stock_code,
                                   exchange_code="NSE",
                                   product_type="cash",
                                   )
spotdata = pd.DataFrame(sdata['Success'])
spotdata.to_excel(option, sheet_name='Spot')
option.save()
spotdata

n=1
m=1
Strike = ["33500","33600","33700","33800","33900","34000","34100","34200","34300","34400","34500","34600","34700","34800",\
          "34900","35000","35100","35200","35300","35400","35500","35600","35700","35800","35900","36000","36100","36200",\
          "36300","36400","36500","36600"]

for s in Strike:

    cdata = breeze.get_historical_data(interval=interval_2,
                                  from_date= from_date,
                                  to_date= to_date,
                                  stock_code=stock_code,
                                  exchange_code="NFO",
                                  product_type="options",
                                  expiry_date=expiry_date,
                                  right="call",
                                  strike_price=s)
    calldata = pd.DataFrame(cdata['Success'])
    print(calldata)
    calldata.to_excel(strikes, sheet_name='Call',startrow=n)
    n+=2000

    pdata = breeze.get_historical_data(interval=interval_2,
                                  from_date= from_date,
                                  to_date= to_date,
                                  stock_code=stock_code,
                                  exchange_code="NFO",
                                  product_type="options",
                                  expiry_date=expiry_date,
                                  right="put",
                                  strike_price=s)
    putdata = pd.DataFrame(pdata['Success'])
    print(putdata)
    putdata.to_excel(strikes, sheet_name='Put',startrow=m )
    m+=2000
    
strikes.save()

def remove(sheet, row):
    
    for cell in row:

        if cell.value != None:
              return
        sheet.delete_rows(row[0].row, 1)

import openpyxl
book = openpyxl.load_workbook("C:/Users/Nish Parikh/Desktop/ipynb Codes/Test.xlsx")

sheet = book['Sheet1']
print(sheet.max_column)
for row in sheet:
    remove(sheet,row)
book.save('Test.xlsx')
exit()
