import webbrowser
from Function.Crawler import Crawler
from Function.Map_CO import Map_CO
from Function.SaveExcel import SaveExcel
from selenium import webdriver

class Main:

    def __init__(self):
        self._Crawler = Crawler()
        self._Map_CO = Map_CO()
        self._SaveExcel = SaveExcel()

    def Main(self):
        self._Crawler.crawler()
        self._SaveExcel.Get_Json()
        self._Map_CO.Read_Excel()
        webbrowser.open_new_tab('render.html')


if __name__ == '__main__':
    Main = Main()
    Main.Main()
