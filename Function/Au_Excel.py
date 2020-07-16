import json
import pandas as pd
import openpyxl as ox


class Au_Excel:
    # Create the list
    confirmlist = []
    activelist = []
    deadlsit = []
    deadRatelist = []
    State_list = []

    def __init__(self):
        pass

    def Get_Json(self):

        with open('Au.json', 'r', encoding='utf-8') as f:
            data = f.read()
        # let the json data change to python data style
        data = json.loads(data)
        # get the China areas
        AuArea = data['sheets']['latest totals']
        update_time = AuArea[1]['Last updated']
        for x in range(len(AuArea)):
            StateList = AuArea[x]['State or territory']
            Death = AuArea[x]['Deaths']
            confirm = AuArea[x]['Confirmed cases (cumulative)']
            Active = AuArea[x]['Active cases']
            deadRate = AuArea[x]['Percent positive']
            State_Co = {'StateList': StateList, 'confirm': confirm, 'Active': Active,
                     'Death': Death, 'deadRate': deadRate}
            self.State_list.append(State_Co)

        CO_In_Au = pd.DataFrame(self.State_list)
        print("Au_In_Chi data list create success")

        # add the data in the lists
        self.confirmlist.append(CO_In_Au['confirm'].values.tolist())
        self.activelist.append(CO_In_Au['Active'].values.tolist())
        self.deadlsit.append(CO_In_Au['Death'].values.tolist())
        self.deadRatelist.append(CO_In_Au['deadRate'].values.tolist())

        print(CO_In_Au)

        # CO_In_Au['confirm'] = self.confirmlist
        # CO_In_Au['Active'] = self.activelist
        # CO_In_Au['Dead'] = self.deadlsit
        # CO_In_Au['deadRate'] = self.deadRatelist

        self.save_excel(CO_In_Au, update_time)

    def save_excel(self, CO_In_Au, update_time):

        # loading the China.xlsx
        writer = pd.ExcelWriter('Au.xlsx', engine='openpyxl')
        try:
            # try to open an existing workbook
            writer.book = ox.load_workbook('Au.xlsx')
            # create and copy the sheet
            writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
            # write out the new sheet
            CO_In_Au.to_excel(writer, sheet_name=update_time[:update_time.find(' ')], index=False)
            # save the data in the xlsx
            writer.save()
            writer.close()

        except FileNotFoundError:
            pass


if __name__ == '__main__':
    Au_Excel =Au_Excel()
    Au_Excel.Get_Json()