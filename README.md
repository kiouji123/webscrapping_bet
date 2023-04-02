# webscrapping_bet
Web scrapping for sport bet actual just for soccer.

The project is composed by 5 files.

1) main_cl_v2.py (You retrieved the data and made the forecast before updating the betdash file, creating the data file and sending the mail.)  Output: data_(date).xlsx
2) res.py (Retrieve the day's result and make the comparison after creating an Excel file and updating the data file to send mail.)  Output: res_(date).xlsx
3) result.bat (Execute the res.py file.)
4) run_script.bat (Execute the main_cl_v2.py file and later turn off the pc.)
5) betdash.xlsx (Recover and consolidate the files from the repository for create a historic data.)

The project go in this order.

4) execute on 01:30
3) execute on 23:00


The project go to a web page and recover fresh data for soccer matches, later createa  a file in a specific path, refresh the data an send the email, to a friends.
PS: Dont forget to change the parameters.


Package            Version
------------------ -----------
async-generator    1.10
attrs              22.2.0
beautifulsoup4     4.11.1
certifi            2022.12.7
cffi               1.15.1
charset-normalizer 2.1.1
colorama           0.4.6
et-xmlfile         1.1.0
h11                0.14.0
idna               3.4
MouseInfo          0.1.3
numpy              1.24.1
openpyxl           3.1.1
outcome            1.2.0
packaging          23.0
pandas             1.5.2
pip                23.0.1
PyAutoGUI          0.9.53
pycparser          2.21
PyGetWindow        0.0.9
PyMsgBox           1.0.9
pyperclip          1.8.2
PyRect             0.2.0
PyScreeze          0.1.28
PySocks            1.7.1
python-dateutil    2.8.2
python-dotenv      0.21.1
pytweening         1.0.4
pytz               2022.7
regex              2022.10.31
requests           2.28.1
selenium           4.8.0
setuptools         65.5.0
six                1.16.0
sniffio            1.3.0
sortedcontainers   2.4.0
soupsieve          2.3.2.post1
tabulate           0.9.0
tqdm               4.64.1
trio               0.22.0
trio-websocket     0.9.2
urllib3            1.26.13
webdriver-manager  3.8.5
wsproto            1.2.0
