import datetime as dt
import pandas as pd
import requests
import numpy as np
from lxml import html
import warnings

# ignore warnings
warnings.filterwarnings('ignore')


class LoadDataWeekChange:

    def __init__(self, source: pd.Series, init: pd.Series, type_spot: pd.Series):
        self.source = source
        self.ytd_df = init
        self.type_spot = type_spot

        self.current = pd.DataFrame()
        self.previous = pd.DataFrame()
        self.change = None

    def load_data(self):
        for name in self.source.index:
            print(f'Loading {name}...')
            self.current[name], self.previous[name] = self.parser(name, self.source.loc[name])
        self.current = self.current.applymap(
            lambda x: x if type(x) == float else float(x) if not '%' in x else float(x[:-1])).T
        self.previous = self.previous.applymap(
            lambda x: x if type(x) == float else float(x) if not '%' in x else float(x[:-1])).T

    @staticmethod
    def parser(item, site):
        """
        Args:
            item: (str) name of the asset
            site: (str) url for scratching

        Returns:
            values of the asset at close of every day in the previous week and the friday before that
        """
        page = requests.get(site, headers={'User-Agent': 'Mozilla/5.0'})
        tree = html.fromstring(page.content)
        week = list()

        if item in ("€STER", "Libor 3M (USD)", "Euribor 3M"):
            data = tree.xpath('//table/tr[@class="tabledata1"]/td/text() | //table/tr[@class="tabledata2"]/td/text()')
            last = data[11].replace("\xa0", "")

            for i in np.arange(1, 11, 2):
                week.insert(0, data[i].replace("\xa0", ""))
        elif item in ("Brent $/bbl", "Gold Spot $/oz"):
            arbre = tree.xpath('//table[@class="cr_dataTable"]/tbody')
            for t in arbre[0]:
                week.insert(1, t.text_content().split()[4].replace("'", ""))
            week = week[1:]
            last = tree.xpath('//span[@id="quote_val"]/text()')
        else:
            arbre = tree.xpath('//table[@class="cr_dataTable"]/tbody')
            for i in range(7):
                week.insert(0, arbre[0][i].text_content().split()[4].replace("'", ""))
            last = [week[1]]
            week = week[2:]

        return week, last

    def run(self):
        self.load_data()
        self.get_df_change()
        return self.current, self.change

    def get_df_change(self):
        """
        Returns:
            weekly and year to date changes
        """
        all_assets = pd.DataFrame()

        all_assets["Last Week"] = self.previous.iloc[:, 0]
        all_assets["This Week"] = self.current.iloc[:, -1]
        all_assets["As of Jan 1st"] = self.ytd_df

        rates = all_assets[self.type_spot.loc[all_assets.index] == 'rate'] / 100
        rates["Weekly Change"] = (rates["This Week"] - rates["Last Week"]) * 10000
        rates["YTD"] = (rates["This Week"] - rates["As of Jan 1st"]) * 10000

        other = all_assets[self.type_spot.loc[all_assets.index] == 'other']
        other["Weekly Change"] = other[["Last Week", "This Week"]].pct_change(axis=1)["This Week"]
        other["YTD"] = (other["This Week"] - other["As of Jan 1st"]) / other["As of Jan 1st"]

        self.change = pd.concat((other, rates))
        self.change = self.change.reindex(['Last Week', 'This Week', 'Weekly Change', 'As of Jan 1st', 'YTD'], axis=1)
        self.change = self.change.reindex(self.source.index, axis=0)


if __name__ == '__main__':
    path = './Sources.xlsx'
    sources = pd.read_excel(path, sheet_name='sources', index_col=0)

    # load and format data
    prices, change = LoadDataWeekChange(sources['Source'], sources['Init'], sources['Type']).run()

    # write results in excel file
    writer = pd.ExcelWriter('./results.xlsx')
    prices.to_excel(writer, sheet_name='prices')
    change.to_excel(writer, sheet_name='changes')
    writer.save()
