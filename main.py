import webbrowser
from google.auth.exceptions import GoogleAuthError
from google.oauth2 import service_account
from googleapiclient.discovery import build
from kivy.clock import Clock
from kivy.core.window import Window
from kivy.lang import Builder
# --------------------------------------------
from kivymd.app import MDApp
from kivymd.uix.screen import MDScreen
from kivymd.uix.screenmanager import ScreenManager
# --------------------------------------------
# GoogleSheet.py

SERVICE_ACCOUNT_FILE = 'keys.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
creds = None
creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# The ID spreadsheet.
SAMPLE_SPREADSHEET_ID = '1vgBQbfhi1MG6GOeOfPu_YNNk7VbRKIa1CJhdCD2FBzE'
service = build('sheets', 'v4', credentials=creds)

# Call the Sheets API
sheet = service.spreadsheets()


class GoogleSheet:

    def __init__(self, username, password):
        self.username = username
        self.password = password

    error_text = None
    name_text = None
    TriOne_List = None
    TriTwo_List = None
    TriThree_List = None

    Dic_of_Result = {}

    def get_name(self, parameter):
        name_cell = 'IDSheet!C' + str(parameter)
        getname = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=name_cell)
        response = getname.execute()
        values = response.get('values', [])
        list_items = values[0][0]
        self.Dic_of_Result.update(Name=list_items)
        return self.get_tri_one_result(parameter)

    def get_tri_one_result(self, parameter):
        name_cell = 'TriOne!B' + str(parameter) + ':F' + str(parameter)
        getname = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=name_cell)
        response = getname.execute()
        values = response.get('values', [])
        self.TriOne_List = list(values[0])
        # check error in cell
        error_value = self.TriOne_List.count('#VALUE!')
        if error_value == 1:
            self.TriOne_List.remove('#VALUE!')
            self.TriOne_List.insert(3, '*')
        self.Dic_of_Result.update(TriOne=self.TriOne_List)
        return self.get_tri_two_result(parameter)

    def get_tri_two_result(self, parameter):
        name_cell = 'TriTwo!B' + str(parameter) + ':F' + str(parameter)
        getname = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=name_cell)
        response = getname.execute()
        values = response.get('values', [])
        self.TriTwo_List = list(values[0])
        # check error in cell
        error_value = self.TriTwo_List.count('#VALUE!')
        if error_value == 1:
            self.TriTwo_List.remove('#VALUE!')
            self.TriTwo_List.insert(3, '*')
        self.Dic_of_Result.update(TriTwo=self.TriTwo_List)
        return self.get_tri_three_result(parameter)

    def get_tri_three_result(self, parameter):
        name_cell = 'TriThree!B' + str(parameter) + ':F' + str(parameter)
        getname = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=name_cell)
        response = getname.execute()
        values = response.get('values', [])
        # check error in cell
        self.TriThree_List = list(values[0])
        error_value = self.TriThree_List.count('#VALUE!')
        if error_value == 1:
            self.TriThree_List.remove('#VALUE!')
            self.TriThree_List.insert(3, '*')
        self.Dic_of_Result.update(TriThree=self.TriThree_List)
        return GoogleSheet.Dic_of_Result

    def operation(self, *args):
        try:
            username_list = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="IDSheet!A3:B250")
            response = username_list.execute()
            values = response.get('values', [])
            is_username = self.username
            is_password = self.password
            b = 0
            for i in range(len(values)):
                if is_username == values[i][0]:
                    if is_password == values[i][1]:
                        a = i + 3
                        self.get_name(a)
                        self.error_text = 'welcome'
                        break
                    else:
                        self.error_text = 'password incorrect'
                        break
                else:
                    b += 1
                if b == len(values):
                    self.error_text = 'username incorrect'
                    break
        except GoogleAuthError:
            self.error_text = 'please check your connexion and try again'
        GoogleSheet.error_text = self.error_text
        return GoogleSheet.error_text

# end of GoogleSheet.py

Window.keyboard_anim_args = {'t': 'in_out_quart', 'd': .2}

Window.softinput_mode = 'below_target'

class Login(MDScreen):
    def on_enter(self, *args):
        Clock.schedule_once(callback=self.show_tri_one_result, timeout=2)

    def show_tri_one_result(self, *args):
        self.ids.label0.text = str(GoogleSheet.Dic_of_Result.get('Name'))
        tri_one_list = list(GoogleSheet.Dic_of_Result.get('TriOne'))
        self.ids.label11.text = tri_one_list[0]
        self.ids.label12.text = tri_one_list[1]
        self.ids.label13.text = tri_one_list[2]
        self.ids.label14.text = tri_one_list[3]
        self.ids.label15.text = tri_one_list[4]
        return Clock.schedule_once(callback=self.show_tri_two_result, timeout=1)

    def show_tri_two_result(self, *args):
        tri_two_list = list(GoogleSheet.Dic_of_Result.get('TriTwo'))
        self.ids.label21.text = tri_two_list[0]
        self.ids.label22.text = tri_two_list[1]
        self.ids.label23.text = tri_two_list[2]
        self.ids.label24.text = tri_two_list[3]
        self.ids.label25.text = tri_two_list[4]
        return Clock.schedule_once(callback=self.show_tri_three_result, timeout=1)

    def show_tri_three_result(self, *args):
        tri_three_list = list(GoogleSheet.Dic_of_Result.get('TriThree'))
        self.ids.label31.text = tri_three_list[0]
        self.ids.label32.text = tri_three_list[1]
        self.ids.label33.text = tri_three_list[2]
        self.ids.label34.text = tri_three_list[3]
        self.ids.label35.text = tri_three_list[4]

    def open_messenger(self):
        webbrowser.open_new('m.me/madanibenslilih')


class Main(MDScreen):
    def on_enter(self, *args):
        self.ids.username_id.text = ''
        self.ids.password_id.text = ''
        self.ids.label.text = ''

    def get_result(self, *args):
        username = self.ids.username_id.text
        password = self.ids.password_id.text
        GoogleSheet(username, password).operation()
        Clock.schedule_once(callback=self.set_result, timeout=2)

    def set_result(self, *args):
        self.ids.label.text = str(GoogleSheet.error_text)
        if self.ids.label.text == 'password incorrect' or self.ids.label.text == 'username incorrect':
            self.ids.label.color = 'red'
        if self.ids.label.text == 'please check your connexion and try again':
            self.ids.label.color = 'blue'
        if self.ids.label.text == 'welcome':
            self.ids.label.color = 'green'
            self.manager.current = 'login'


Builder.load_file('main.kv')


class DesignApp(MDApp):

    def build(self):
        sm = ScreenManager()
        sm.add_widget(Main(name='main'))
        sm.add_widget(Login(name='login'))
        return sm


if __name__ == '__main__':
    DesignApp().run()
