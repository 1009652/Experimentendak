from flask import Flask,request,json
import xlwt
from xlwt import Workbook
from datetime import datetime

app = Flask(__name__)

wb = Workbook()
sheet1 = wb.add_sheet('Sound Data')
amountOfEntries = 0
sheet1.write(amountOfEntries, 0, "Date Time:")
sheet1.write(amountOfEntries, 1, "Sound Level")
amountOfEntries += 1
wb.save('Sound_Data.xls')


@app.route('/')
def hello():
    return 'Webhooks with Python'

@app.route('/uplink',methods=['POST'])
def getMicrofoonData():
    global amountOfEntries
    data = request.json
    strippedData = data['uplink_message']['decoded_payload']['data']
    now = datetime.now()
    sheet1.write(amountOfEntries, 0, now.strftime("%Y-%m-%d %H:%M"))
    sheet1.write(amountOfEntries, 1, strippedData + " dB")
    amountOfEntries += 1
    try:
        wb.save('Sound_Data.xls')
    except:
        print("Can't have excel file opened while writing data, close file to continue")
    return data
 
if __name__ == '__main__':
    app.run(debug=True)