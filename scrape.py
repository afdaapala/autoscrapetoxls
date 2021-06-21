import requests
from bs4 import BeautifulSoup as bs
import urllib3
import win32com.client
from pathlib import Path

#path 
f_path = Path.cwd()
f_name = 'xl.xlsm'
filename = f_path/f_name
sheetname = 'pracu'
#buka excel
excel_app = win32com.client.gencache.EnsureDispatch("Excel.Application")
#buka workshet
workst = excel_app.Workbooks.Open(filename)

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
url = "https://data.bmkg.go.id/DataMKG/MEWS/DigitalForecast/DigitalForecast-SulawesiTenggara.xml"
response = requests.get(url,verify=False)
r = response.text

cont = bs(r,"xml")
# print(cont)
tanggal = {'hari ini' : '', 'besok':''}
tempat = {'baubau': '', 'kendari': '' }

kode = {
'0': 'Cerah',
'1': 'Cerah Berawan',
'2': 'Cerah Berawan',
'3': 'Berawan',
'4': 'Berawan Tebal',
'5': 'Udara Kabur',
'10': 'Asap',
'45': 'Kabut',
'60': 'Hujan Ringan',
'61': 'Hujan Sedang',
'63': 'Hujan Lebat',
'80': 'Hujan Lokal',
'95': 'Hujan Lebat Disertai Petir',
'97': 'Hujan Lebat Disertai Petir'
}

#fetchdate
tyear = cont.find("issue").year.string
tmonth = cont.find("issue").month.string
tday = cont.find("issue").day.string
tanggal['hari ini'] = tyear + tmonth + tday
hariini = tanggal['hari ini']

workst.Sheets(sheetname).Range("B4").Value = hariini

# Cuaca siang / kota
baubau = cont.find(id="501512").find(id="weather").find(h="6").value.string
baubauhmax = cont.find(id="501512").find(id="humax").find(day=hariini).value.string
baubautmax = cont.find(id="501512").find(id="tmax").find(day=hariini).value.string
baubauhmin = cont.find(id="501512").find(id="humin").find(day=hariini).value.string
baubautmin = cont.find(id="501512").find(id="tmin").find(day=hariini).value.string

kendari = cont.find(id="501513").find(id="weather").find(h="6").value.string
kendarihmax = cont.find(id="501513").find(id="humax").find(day=hariini).value.string
kendaritmax = cont.find(id="501513").find(id="tmax").find(day=hariini).value.string
kendarihmin = cont.find(id="501513").find(id="humin").find(day=hariini).value.string
kendaritmin = cont.find(id="501513").find(id="tmin").find(day=hariini).value.string

#just in case
tempat['baubau'] = kode[baubau]
tempat['kendari'] = kode[kendari]

#Insert ke cell
workst.Sheets(sheetname).Range("C6").Value = baubautmin
workst.Sheets(sheetname).Range("D6").Value = baubautmax
workst.Sheets(sheetname).Range("E6").Value = baubauhmin
workst.Sheets(sheetname).Range("F6").Value = baubauhmax
workst.Sheets(sheetname).Range("G6").Value = tempat['baubau']

workst.Sheets(sheetname).Range("C7").Value = kendaritmin
workst.Sheets(sheetname).Range("D7").Value = kendaritmax
workst.Sheets(sheetname).Range("E7").Value = kendarihmin
workst.Sheets(sheetname).Range("F7").Value = kendarihmax
workst.Sheets(sheetname).Range("G7").Value = tempat['kendari']