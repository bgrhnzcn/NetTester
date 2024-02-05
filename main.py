import datetime as dt
from socket import timeout
from urllib import request
from urllib import error
import xlsxwriter as xs
import keyboard


class NetTester:
    index: int
    timeZone: dt.timezone
    workbook: xs.Workbook

    def __init__(self, timezone_offset: int, output_name: str):
        self.index = 0
        self.timeZone = dt.timezone(dt.timedelta(hours=timezone_offset))
        self.workbook = xs.Workbook(output_name)

    @staticmethod
    def internet_check(lst):
        try:
            try:
                request.urlopen("https://www.google.com", timeout=1)
                return True
            except error.URLError:
                return False
        except timeout:
            return lst

    def check_worksheet(self):
        ws = self.workbook.get_worksheet_by_name(dt.date.today().strftime("%d-%m-%Y"))
        if ws is None:
            self.index = 0
            return self.workbook.add_worksheet(name=dt.date.today().strftime("%d-%m-%Y"))
        else:
            return ws


tester = NetTester(3, "data.xlsx")
last_state = False
while True:
    worksheet = tester.check_worksheet()
    current_state = NetTester.internet_check(last_state)
    if current_state is not last_state:
        last_state = current_state
        if current_state is True:
            current_time = dt.datetime.now(tester.timeZone)
            worksheet.write(tester.index, 0, current_time.time().strftime("%H:%M:%S"))
            worksheet.write(tester.index, 1, "Connected")
        else:
            current_time = dt.datetime.now(tester.timeZone)
            worksheet.write(tester.index, 0, current_time.time().strftime("%H:%M:%S"))
            worksheet.write(tester.index, 1, "Disconnected")
        tester.index += 1
    if keyboard.is_pressed('q'):
        tester.workbook.close()
        break
