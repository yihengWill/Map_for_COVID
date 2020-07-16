import json
import pandas as pd
import openpyxl as ox


class SaveExcel:
    # Create the list
    confirmlist = []
    suspectlist = []
    deadlsit = []
    heallist = []
    deadRatelist = []
    healRatelist = []
    city_list = []

    def __init__(self):
        pass

    def Get_Json(self):

        # Read the json file
        with open('China.json', 'r', encoding='utf-8') as f:
            data = f.read()
        # let the json data change to python data style
        data = json.loads(data)
        # get the China areas
        ChinaArea = data['areaTree'][0]
        # get update time
        ReportTimer = data['lastUpdateTime']

        # get the province detail
        provinceList = ChinaArea['children']

        for x in range(len(provinceList)):
            province = provinceList[x]['name']
            province_list = provinceList[x]['children']
            for y in range(len(province_list)):
                city = province_list[y]['name']
                total = province_list[y]['total']
                today = province_list[y]['today']
                china_dict = {'province': province, 'city': city, 'total': total,
                              'today': today}
                self.city_list.append(china_dict)

        # Create the China COVID-19 data for using dataframe
        CO_In_Chi = pd.DataFrame(self.city_list)
        print("Total data list create success")

        # let the data to list
        for value in CO_In_Chi['total'].values.tolist():
            # add the data in the lists
            self.confirmlist.append(value['confirm'])
            self.suspectlist.append(value['suspect'])
            self.deadlsit.append(value['dead'])
            self.heallist.append(value['heal'])
            self.deadRatelist.append(value['deadRate'])
            self.healRatelist.append(value['healRate'])

        # add the add in different sheet
        CO_In_Chi['confirm'] = self.confirmlist
        CO_In_Chi['suspect'] = self.suspectlist
        CO_In_Chi['dead'] = self.deadlsit
        CO_In_Chi['heal'] = self.heallist
        CO_In_Chi['deadRate'] = self.deadRatelist
        CO_In_Chi['healRate'] = self.healRatelist

        # Create today data for China
        today_confirmlist = []
        today_confirmCutslist = []
        # let the data to list
        for value in CO_In_Chi['today'].values.tolist():
            # add the data in the lists
            today_confirmlist.append(value['confirm'])
            today_confirmCutslist.append(value['confirmCuts'])

        CO_In_Chi['today_confirm'] = today_confirmlist
        CO_In_Chi['today_confirmCuts'] = today_confirmCutslist

        # delete total and today list
        CO_In_Chi.drop(['total', 'today'], axis=1, inplace=True)

        # Call the save excel function
        self.save_excel(CO_In_Chi, ReportTimer)

    def save_excel(self, CO_In_Chi, ReportTimer):

        # loading the China.xlsx
        writer = pd.ExcelWriter('China.xlsx', engine='openpyxl')
        try:
            # try to open an existing workbook
            writer.book = ox.load_workbook('China.xlsx')
            # create and copy the sheet
            writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
            # write out the new sheet
            CO_In_Chi.to_excel(writer, sheet_name=ReportTimer[:ReportTimer.find(' ')], index=False)
            # save the data in the xlsx
            writer.save()
            writer.close()

        except FileNotFoundError:
            pass

