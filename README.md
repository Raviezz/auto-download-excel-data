# auto-download-excel-data
.@Author Raviteja
 >Download content using urls from a excel sheet 

1. First install python 3 locally
2. Then update python with requirement.txt 
  >pip install -r requirements.txt
3. Update the XLS file (downloaded from mail) URL in the __main__
4. Update the DEFAULT SCREENSHOTS URL in the __main__
5. Updat your excel shhet name Ex: practiceplusapps5_Set4
6. Dont worry about downloads folder , I defined it as "D:\\DRL ORG\\' + str(time.datetime.today())[:10]" current date
7.Run read_xls_file.py only
  >python read_xls_file.py
