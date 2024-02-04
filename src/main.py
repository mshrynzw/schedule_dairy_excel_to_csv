# -*- coding: utf-8 -*-
from common import set_log, set_common
from logging import getLogger, config
from datetime import datetime, timedelta

import openpyxl as excel

log_conf = set_log()
config.dictConfig(log_conf)
logger = getLogger(__name__)

common_conf = set_common()

if __name__ == '__main__':
    logger.info("Start")

    book = excel.load_workbook("./excel/スケジュール.xlsx", data_only=True)
    csv_body = "authorId,datetime,scheduleMemo,scheduleId,dairyId,dairyMemo,mentalId\n"
    for sheet in book.worksheets:
        if sheet.title == "Todo" or sheet.title == "グラフ" or sheet.title == "conf":
            continue

        sheet_name = sheet.title
        date_obj = datetime.strptime(sheet_name, "%y%m%d")

        days = []
        for i in range(0, 7):
            date_obj_before = date_obj + timedelta(days=i)
            date_str_before = date_obj_before.strftime("%y%m%d")
            days.append(date_str_before)

        col = [
            ["A", "C", "D", "E"],
            ["A", "H", "I", "J"],
            ["A", "M", "N", "O"],
            ["A", "R", "S", "T"],
            ["A", "W", "X", "Y"],
            ["A", "AB", "AC", "AD"],
            ["A", "AG", "AH", "AI"]
        ]
        for i, c in enumerate(col):
            for row in range(5, 53):
                time = sheet[c[0] + str(row)].value.strftime("%H%M")
                schedule = sheet[c[1] + str(row)].value
                schedule = "" if schedule is None else schedule
                dairy = sheet[c[2] + str(row)].value
                dairy = "" if dairy is None else dairy
                mental = sheet[c[3] + str(row)].value
                mental = "" if mental is None else mental

                if schedule != "":
                    schedule = common_conf["schedule_dairy"][schedule]

                if dairy != "":
                    dairy = common_conf["schedule_dairy"][dairy]

                if mental != "":
                    mental = common_conf["mental"][str(mental)]

                dt = days[i] + time
                csv_body += ",20{},,{},{},,{}\n".format(dt, schedule, dairy, mental)

    with open("./csv/import_schedule_dairy.csv", "w", newline="") as file:
        file.write(csv_body)

    logger.info("End")
