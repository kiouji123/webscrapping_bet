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
