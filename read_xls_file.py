"""
pip install -r requirements.txt
"""
import pandas as pd
import os
import requests
import datetime as time


class ReadExcel:

    def read_data_from_excel(self, folder_name, xls_folder, sheet_name, start, end):
        try:
            xlsx_data = pd.read_excel(xls_folder, sheet_name=sheet_name, keep_default_na=False)
            xlsx_data = xlsx_data[start:end]
            for index, row in xlsx_data.iterrows():
                print("Index", index)
                path = folder_name + '\\' + row['Title']
                print(path)
                if not os.path.exists(path):
                    os.mkdir(path)
                apk_file = requests.get(row['Android APK Link'])
                image_512_512 = requests.get(row['Graphic assets(512*512)'])
                image_512_1024 = requests.get(row['Feature graphic(1024*500)'])
                consent_form = requests.get(row['Consent Form'])
                open(path + '\\' + row['Title'] + '.apk', 'wb').write(apk_file.content)
                print(row['Android APK Link'])
                open(path + '\\' + row['Title'] + '512_512.png', 'wb').write(image_512_512.content)
                print(row['Graphic assets(512*512)'])
                open(path + '\\' + row['Title'] + '1024_512.png', 'wb').write(image_512_1024.content)
                print(row['Feature graphic(1024*500)'])
                open(path + '\\' + row['Title'] + '.pdf', 'wb').write(consent_form.content)
                print(row['Consent Form'])
                file = open(path + '\\doc_details.txt', 'w+')
                file.write('App Title:\n\n')
                file.writelines(row['Title'] + '\n\n\n')
                file.write('Full Description:\n\n')
                file.writelines(row['Full description'] + '\n\n\n')
                file.write('Short description:\n\n')
                file.writelines(row['Short description'] + '\n\n\n')
                file.write('Email :\n\n')
                file.writelines(row['Email'] + '\n\n\n')
                file.write('Privacy Policy :\n\n')
                file.writelines(row['Privacy Policy'] + '\n\n\n')
                file.write('Content Rating Email :\n\n')
                file.writelines(row['Content Rating Email'] + '\n\n\n')
                file.write('Additional Information :\n\n')
                file.writelines(row['Additional Information'])
        except FileNotFoundError:
            print("Something is wrong , check")
        finally:
            file.close()


if __name__ == "__main__":
    excel = ReadExcel()
    dir_name = 'D:\\DRL ORG\\' + str(time.datetime.today())[:10]
    os.mkdir(dir_name)
    excel_folder = 'D:\\DRL ORG\\Downloads\\s3_upload_data_OnlyDRL_Set4.xls'
    sheet_name = 'practiceplusapps5_Set4'
    start_from = 6
    end_till = 10
    excel.read_data_from_excel(dir_name, excel_folder, sheet_name, start_from, end_till)
