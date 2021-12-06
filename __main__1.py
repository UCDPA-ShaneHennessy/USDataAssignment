import xlrd
from prettytable import PrettyTable

import matplotlib.pyplot as plt
import numpy as np


class ShaneHennessyUSData:

    if __name__ == "__main__":

        # Name of the file
        loc = "US Data Python Project.xlsx"

        # Open the excel file
        wb = xlrd.open_workbook(loc)

        sheet = wb.sheet_by_index(0)

        sheet3 = PrettyTable(["Type", "Storage", "In Service", "Total", "% of the Market"])

        data = {}

        for i in range(1, sheet.nrows):
            valueType = str(sheet.cell_value(i, 0))

            if valueType == "737 (CFMI)" or valueType == "737 (JT8D)" or valueType == "737 NG" or valueType == "A319" or valueType == "A320" or valueType == "A321":
                # data[valueType + "Storage"] = 0
                # data[valueType + "InService"] = 0
                data[valueType] = [0, 0, 0, 0.0]

        data["Total"] = [0, 0, 0, 0.0]

        counter = 0

        for i in range(1, sheet.nrows):
            valueType = str(sheet.cell_value(i, 0))
            status = str(sheet.cell_value(i, 4))

            if valueType == "737 (CFMI)" or valueType == "737 (JT8D)" or valueType == "737 NG" or valueType == "A319" or valueType == "A320" or valueType == "A321":

                if str(status) == "Storage":
                    data[valueType][0] += 1
                elif str(status) == "In Service":
                    data[valueType][1] +=1

        for i in data:
            # print(i)
            data[i][2] = data[i][0] + data[i][1]

        for i in data:
            if i is not "Total":
                data["Total"][0] += data[i][0]
                data["Total"][1] += data[i][1]
                data["Total"][2] += data[i][2]

        for i in data:
            data[i][3] = round(data[i][2] / data["Total"][2] * 100, 1)

        sheet3.add_row(["737 (CFMI)", data["737 (CFMI)"][0], data["737 (CFMI)"][1], data["737 (CFMI)"][2], data["737 (CFMI)"][3]])
        sheet3.add_row(
            ["737 (JT8D)", data["737 (JT8D)"][0], data["737 (JT8D)"][1], data["737 (JT8D)"][2], data["737 (JT8D)"][3]])
        sheet3.add_row(
            ["737 NG", data["737 NG"][0], data["737 NG"][1], data["737 NG"][2], data["737 NG"][3]])
        sheet3.add_row(
            ["A319", data["A319"][0], data["A319"][1], data["A319"][2], data["A319"][3]])
        sheet3.add_row(
            ["A320", data["A320"][0], data["A320"][1], data["A320"][2], data["A320"][3]])
        sheet3.add_row(
            ["A321", data["A321"][0], data["A321"][1], data["A321"][2], data["A321"][3]])
        sheet3.add_row(["", "", "", "", ""])
        sheet3.add_row(
            ["Total", data["Total"][0], data["Total"][1], data["Total"][2], data["Total"][3]])

        print(sheet3)

        sheet3_T2 = PrettyTable(["Type", "Storage", "In Service", "Total", "% of the Market"])

        data["737"] = [0, 0, 0, 0.0]
        data["AA320"] = [0, 0, 0, 0.0]
        data["Average Total"] = [0, 0, 0, 0.0]

        for i in data:
            if i not in ["Total", "737", "AA320"]:
                if i[0] == "A":
                    data["AA320"][0] += data[i][0]
                    data["AA320"][1] += data[i][1]
                else:
                    data["737"][0] += data[i][0]
                    data["737"][1] += data[i][1]

        for i in data:
            if i in ["737", "AA320"]:
                data[i][2] = data[i][0] + data[i][1]

        for i in data:
            if i in ["737", "AA320"]:
                data["Average Total"][0] += data[i][0]
                data["Average Total"][1] += data[i][1]
                data["Average Total"][2] += data[i][2]

        for i in data:
            if i in ["737", "AA320", "Average Total"]:
                data[i][3] = round(data[i][2] / data["Total"][2] * 100)

        sheet3_T2.add_row(["737", data["737"][0], data["737"][1], data["737"][2], data["737"][3]])
        sheet3_T2.add_row(
            ["A320", data["AA320"][0], data["AA320"][1], data["AA320"][2], data["AA320"][3]])
        sheet3_T2.add_row(["", "", "", "", ""])
        sheet3_T2.add_row(
            ["Average Total", data["Average Total"][0], data["Average Total"][1], data["Average Total"][2], data["Average Total"][3]])

        print(sheet3_T2)

        # for key, value in data.items():
        #     print(key, value)

        plotPie = []
        plotPieLabel = []
        for i in data:
            if i not in ["Total", "737", "AA320", "Average Total"]:
                plotPie.append(data[i][3])
                plotPieLabel.append(i)

        plt.pie(np.array(plotPie), labels=plotPieLabel)
        plt.show()