# auto-download-excel-data
@Author Raviteja
Download content using urls from a excel sheet 

1. First install python 3 locally
2. Then update python with requirement.txt 
  >pip install -r requirements.txt
3. Update the DEFAULT SCREENSHOTS URL in the __main__
4. Updat your excel shhet name Ex: practiceplusapps5_Set4
5. Dont worry about downloads folder , I defined it as "D:\\DRL ORG\\' + str(time.datetime.today())[:10]" // current data
6.Run read_xls_file.py only
  >python read_xls_file.py
