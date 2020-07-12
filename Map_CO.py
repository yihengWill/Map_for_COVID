import pandas as pd
from pyecharts import options as opts
from pyecharts.charts import Map


class Map_CO:

    def Read_Excel(self):
        # read the excel
        excel = pd.read_excel('China.xlsx', sheet_name='2020-07-12')
        # show the sum of province data
        data = excel.groupby(by='province', as_index=False).sum()
        self.Show_Map(data)

    def Show_Map(self, data):
        # show the Map of China
        Map_list = list(zip(data['province'].values.tolist(), data['confirm'].values.tolist()))

        def map_china() -> Map:
            c = (
                Map()
                    .add(series_name="确诊病例", data_pair=Map_list, maptype="china")
                    .set_global_opts(
                    title_opts=opts.TitleOpts(title="中国疫情实时地图"),
                    visualmap_opts=opts.VisualMapOpts(is_piecewise=True, pieces=[{"max": 30, "min": 0,
                                                "label": "0-9",
                                                "color": "#FFE4E1"},
                                               {"max": 99,
                                                "min": 10,
                                                "label": "10-99",
                                                "color": "#FF7F50"},
                                               {"max": 499,
                                                "min": 100,
                                                "label": "100-499",
                                                "color": "#F08080"},
                                               {"max": 999,
                                                "min": 500,
                                                "label": "500-999",
                                                "color": "#CD5C5C"},
                                               {"max": 9999,
                                                "min": 1000,
                                                "label": "1000-9999",
                                                "color": "#990000"},
                                               {"max": 99999,
                                                "min": 10000,
                                                "label": ">=10000",
                                                "color": "#660000"}]
                                       )
                ))
            return c
        d_map = map_china()
        d_map.render()

if __name__ == '__main__':
    Map_CO = Map_CO()
    Map_CO.Read_Excel()
