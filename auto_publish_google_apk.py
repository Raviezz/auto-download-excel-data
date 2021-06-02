import pyautogui as pa
import clipboard as cp
import pandas as pd
import time as td
import re


class DRLAutomation:

    def create_application_apk(self, file_search, image_path, data, index):

        interval_seconds = 1.0
        td.sleep(1)
        pa.click(x=554, y=748, clicks=1, interval=interval_seconds, button='left')
        X, Y = pa.locateCenterOnScreen(image_path + "\\create_app.png")
        td.sleep(0.5)
        pa.click(X, Y, clicks=1, interval=interval_seconds, button='left')
        X, Y = pa.locateCenterOnScreen(image_path + "\\app_title_name.png")
        td.sleep(0.5)
        pa.click(X, Y, clicks=1, interval=interval_seconds, button='left')
        pa.typewrite(data['Title'][index])
        X, Y = pa.locateCenterOnScreen(image_path + "\\generate_application.png")
        td.sleep(2)
        pa.click(X, Y, clicks=1, interval=interval_seconds, button='left')
        td.sleep(5)
        X, Y = pa.locateCenterOnScreen(image_path + "\\click_app_release_btn.png")
        td.sleep(1)
        pa.click(X, Y, clicks=1, interval=interval_seconds, button='left')
        td.sleep(5)
        pa.hotkey('pgdn')
        td.sleep(3)
        X, Y = pa.locateCenterOnScreen(image_path + "\\release_manage_btn.png")
        td.sleep(0.5)
        pa.click(X, Y, clicks=1, interval=interval_seconds, button='left')
        td.sleep(5)
        pa.hotkey('pgdn')
        td.sleep(3)
        X, Y = pa.locateCenterOnScreen(image_path + "\\create_release_button.png")
        td.sleep(0.5)
        pa.click(X, Y, clicks=1, interval=interval_seconds, button='left')
        td.sleep(5)
        pa.hotkey('pgdn')
        td.sleep(3)
        X, Y = pa.locateCenterOnScreen(image_path + "\\release_continue_mini.png")
        td.sleep(0.5)
        pa.click(X, Y, clicks=1, interval=interval_seconds, button='left')
        td.sleep(8)
        X, Y = pa.locateCenterOnScreen(image_path + "\\browse_apk_btn.png")
        td.sleep(0.5)
        pa.click(X, Y, clicks=1, interval=interval_seconds, button='left')
        td.sleep(3)
        X, Y = pa.locateCenterOnScreen(image_path + "\\apk_search_icon.png")
        td.sleep(0.5)
        pa.click(X, Y, clicks=1, interval=interval_seconds, button='left')
        td.sleep(5)
        pa.typewrite(file_search + '\\' + data['Title'][index])
        pa.hotkey('enter')
        td.sleep(3)
        X, Y = pa.locateCenterOnScreen(image_path + "\\select_the_apk_file.png")
        td.sleep(0.5)
        pa.click(X, Y, clicks=1, interval=interval_seconds, button='left')
        td.sleep(3)
        X, Y = pa.locateCenterOnScreen(image_path + "\\open_selected_apk.png")
        td.sleep(0.5)
        pa.click(X, Y, clicks=1, interval=interval_seconds, button='left')
        td.sleep(10)
        pa.hotkey('pgdn', 'pgdn', 'pgdn')
        td.sleep(3)
        X, Y = pa.locateCenterOnScreen(image_path + "\\enter_to_paste_text.png")
        td.sleep(0.5)
        pa.click(X, Y, clicks=1, interval=interval_seconds, button='left')
        td.sleep(1)
        pa.hotkey('ctrl', 'a')
        pa.hotkey('ctrl', 'x')
        desc = '<en-US>\n' + data['Full description'][index] + '\n</en-US>'
        pa.typewrite(desc)
        td.sleep(3)
        # pa.hotkey('pgdn','pgdn')
        X, Y = pa.locateCenterOnScreen(image_path + "\\save_apk_release.png")
        td.sleep(0.5)
        pa.click(X, Y, clicks=1, interval=interval_seconds, button='left')


    def store_listing(self, file_search, image_path, data, index):
        interval_seconds = 1.0
        td.sleep(1)
        # X, Y = pa.locateCenterOnScreen("D:\\pyauto\\chrometab_mini.png")
        td.sleep(1)
        pa.click(x=554, y=748, clicks=1, interval=interval_seconds, button='left')
        td.sleep(1)
        #X, Y = pa.locateCenterOnScreen(image_path + "\\store_list_icon.png")
        #print("store_listing",X,Y)
        td.sleep(1)
        pa.click(43 ,496, clicks=1, interval=interval_seconds, button='left')
        td.sleep(1)
        #X, Y = pa.locateCenterOnScreen(image_path + "\\short_description_text.png")
        #print("short desc", X, Y)
        td.sleep(1)
        pa.click(1269, 536, clicks=1, interval=interval_seconds, button='left')
        pa.typewrite(data['Short description'][index])
        td.sleep(1)
        pa.hotkey('pgdn')
        td.sleep(1)
        #X, Y = pa.locateCenterOnScreen(image_path + "\\full_description_text.png")
        #print("full desc", X, Y)
        td.sleep(2)
        pa.click(1268, 300, clicks=1, interval=interval_seconds, button='left')
        pa.typewrite(data['Full description'][index])
        td.sleep(2)
        #X, Y = pa.locateCenterOnScreen(image_path + "\\whitecoats_com.png")
        #print("whitecoats url", X, Y)
        pa.click(739 ,50, clicks=1, interval=interval_seconds, button='left')
        pa.hotkey('ctrl', 'c')
        url = cp.paste()
        url = self.get_package_name(url)
        print("URL:",url)
        td.sleep(2)
        #pa.click(X, Y, clicks=1, interval=interval_seconds, button='left')
        #X, Y = pa.locateCenterOnScreen(image_path + "\\consent_form_url.png")
        #print("consent form url", X, Y)
        td.sleep(2)
        pa.click(880, 406, clicks=1, interval=interval_seconds, button='left')
        td.sleep(5)
        X, Y = pa.locateCenterOnScreen(image_path + "\\consent_modal.png")
        print("consent form modal", X, Y)
        pa.click(699, 263, clicks=2, interval=interval_seconds, button='left')
        pa.hotkey('pgdn', 'pgdn')
        td.sleep(3)
        X, Y = pa.locateCenterOnScreen(image_path + "\\save_consent_document.png")
        print("save consent form", X, Y)
        pa.click(696, 562, clicks=1, interval=interval_seconds, button='left')
        td.sleep(5)
        X, Y = pa.locateCenterOnScreen(image_path + "\\consent_modal.png")
        print("consent form modal", X, Y)
        pa.click(699, 263, clicks=2, interval=interval_seconds, button='left')
        pa.hotkey('pgdn')
        td.sleep(1)
        pa.hotkey('pgdn')
        td.sleep(1)
        pa.hotkey('pgdn')
        td.sleep(2)
        #X, Y = pa.locateCenterOnScreen(image_path + "\\package_name_apk.png")
        #print("package name", X, Y)
        pa.click(689,594, clicks=1, interval=interval_seconds, button='left')
        td.sleep(3)
        pa.typewrite(url)
        td.sleep(3)
        pa.hotkey('pgdn')
        td.sleep(2)
        X, Y = pa.locateCenterOnScreen(image_path + "\\arrow_down.png")
        pa.click(X, Y, clicks=1, interval=interval_seconds, button='left')
        td.sleep(3)
        #
        X, Y = pa.locateCenterOnScreen(image_path + "\\intellectual_select.png")
        pa.click(X, Y, clicks=1, interval=interval_seconds, button='left')
        td.sleep(3)
        X, Y = pa.locateCenterOnScreen(image_path + "\\choose_consent_form.png")
        pa.click(X, Y, clicks=1, interval=interval_seconds, button='left')
        td.sleep(3)
        X, Y = pa.locateCenterOnScreen(image_path + "\\apk_search_icon.png")
        td.sleep(0.5)
        pa.click(X, Y, clicks=1, interval=interval_seconds, button='left')
        td.sleep(2)
        pa.typewrite(file_search + '\\' + data['Title'][index])
        pa.hotkey('enter')
        td.sleep(3)
        X, Y = pa.locateCenterOnScreen(image_path + "\\select_pdf.png")
        td.sleep(0.5)
        pa.click(X, Y, clicks=1, interval=interval_seconds, button='left')
        td.sleep(5)
        X, Y = pa.locateCenterOnScreen(image_path + "\\open_selected_apk.png")
        td.sleep(0.5)
        pa.click(X, Y, clicks=1, interval=interval_seconds, button='left')
        td.sleep(5)
        pa.hotkey('pgdn')
        X, Y = pa.locateCenterOnScreen(image_path + "\\addtional_information_consent.png")
        td.sleep(0.5)
        pa.click(X, Y, clicks=1, interval=interval_seconds, button='left')
        td.sleep(2)
        pa.typewrite(data['Additional Information'])
        X, Y = pa.locateCenterOnScreen(image_path + "\\uncheck_modal.png")
        td.sleep(0.5)
        pa.click(X, Y, clicks=1, interval=interval_seconds, button='left')
        td.sleep(2)
        X, Y = pa.locateCenterOnScreen(image_path + "\\submit.png")
        td.sleep(0.5)
        pa.click(X, Y, clicks=1, interval=interval_seconds, button='left')

    def read_data_from_excel(self, xls_folder, sheet_name):
        xlsx_data = pd.read_excel(xls_folder, sheet_name=sheet_name, keep_default_na=False)
        return xlsx_data

    def get_package_name(self, url):
        m = re.search('com.whitecoats.patientplus(.+?)appid', url)
        found = ''
        if m:
            found = m.group(1)
        found = str(found)
        print(found)
        return 'com.whitecoats.patientplus' + found[:len(found) - 1]


if __name__ == "__main__":
    drl = DRLAutomation()
    default_dir = 'D:\\DRL ORG\\2020-05-14'
    google_play_images = 'D:\\pyauto'
    excel_folder = 'D:\\DRL ORG\\Downloads\\Set_8.xls'
    sheet_name = 'practiceplusapps5_Set8'
    exec_data = drl.read_data_from_excel(excel_folder, sheet_name)
    # print(exec_data)
    td.sleep(1)
    #print(pa.locateCenterOnScreen("D:\\pyauto\\chrometab_mini_new.png",grayscale=True))

    #td.sleep(3)
    #pa.click(X, Y, clicks=1, button='left')
    #pa.hotkey('pgdn')
    doc_no = 9
    #td.sleep(1)
    #pa.click(x=554, y=748, clicks=1, interval=0.5, button='left')
    #td.sleep(2)
    #pa.click(739,50, clicks=1, interval=0.5, button='left')
    #pa.moveTo(739,50)

    #pa.hotkey('ctrl', 'c')
    #url = cp.paste()
    #print("copid",url)
    #url = drl.get_package_name(url)
    #print("URL:", url)
    #drl.get_package_name()
    drl.create_application_apk(default_dir, google_play_images, exec_data, doc_no)
    #drl.store_listing(default_dir, google_play_images, exec_data, doc_no)