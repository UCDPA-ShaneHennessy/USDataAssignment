import xlrd
from prettytable import PrettyTable

import matplotlib.pyplot as plt
import numpy as np


class ShaneHennessyUSData:

    # This is the main method.
    if __name__ == "__main__":

        # Name of the file
        loc = "US Data Python Project.xlsx"

        # Open the excel file
        wb = xlrd.open_workbook(loc)

        # Get the particular sheet.
        sheet = wb.sheet_by_index(0)

        # Name of the column for the table.
        sheet3 = PrettyTable(["Type", "Storage", "In Service", "Total", "% of the Market"])

        # Variable to store the data.
        data = {}

        # Loop to initialise the variable with default value.
        for i in range(1, sheet.nrows):
            valueType = str(sheet.cell_value(i, 0))

            if valueType == "737 (CFMI)" or valueType == "737 (JT8D)" or valueType == "737 NG" or valueType == "A319" or valueType == "A320" or valueType == "A321":
                # data[valueType + "Storage"] = 0
                # data[valueType + "InService"] = 0
                data[valueType] = [0, 0, 0, 0.0]

        # Create and initialise the total with default values.
        data["Total"] = [0, 0, 0, 0.0]

        counter = 0

        # Loop to do the calculation.
        for i in range(1, sheet.nrows):
            # Get the type.
            valueType = str(sheet.cell_value(i, 0))
            # Get the status.
            status = str(sheet.cell_value(i, 4))

            # Check if the type is our value.
            if valueType == "737 (CFMI)" or valueType == "737 (JT8D)" or valueType == "737 NG" or valueType == "A319" or valueType == "A320" or valueType == "A321":

                # Check if the status is Storage or Service.
                if str(status) == "Storage":
                    data[valueType][0] += 1
                elif str(status) == "In Service":
                    data[valueType][1] +=1

        # Calculation done for the third column.
        for i in data:
            # print(i)
            data[i][2] = data[i][0] + data[i][1]

        # Loop to calculate the total.
        for i in data:
            if i is not "Total":
                data["Total"][0] += data[i][0]
                data["Total"][1] += data[i][1]
                data["Total"][2] += data[i][2]

        # Loop to calculate the forth column and round up the value.
        for i in data:
            data[i][3] = round(data[i][2] / data["Total"][2] * 100, 1)

        # Add all the data to the table.
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

        # Print the table.
        print(sheet3)

        # Second table.
        sheet3_T2 = PrettyTable(["Type", "Storage", "In Service", "Total", "% of the Market"])

        # Extra variables needed for second table.
        data["737"] = [0, 0, 0, 0.0]
        data["AA320"] = [0, 0, 0, 0.0]
        data["Average Total"] = [0, 0, 0, 0.0]

        # Loop to calculate those extra variables.
        for i in data:
            if i not in ["Total", "737", "AA320"]:
                if i[0] == "A":
                    data["AA320"][0] += data[i][0]
                    data["AA320"][1] += data[i][1]
                else:
                    data["737"][0] += data[i][0]
                    data["737"][1] += data[i][1]

        # Loop to calculate the third column
        for i in data:
            if i in ["737", "AA320"]:
                data[i][2] = data[i][0] + data[i][1]

        # Loop to calculate the average total.
        for i in data:
            if i in ["737", "AA320"]:
                data["Average Total"][0] += data[i][0]
                data["Average Total"][1] += data[i][1]
                data["Average Total"][2] += data[i][2]

        # Loop to calculate forth column.
        for i in data:
            if i in ["737", "AA320", "Average Total"]:
                data[i][3] = round(data[i][2] / data["Total"][2] * 100)

        # Add all the data to the table.
        sheet3_T2.add_row(["737", data["737"][0], data["737"][1], data["737"][2], data["737"][3]])
        sheet3_T2.add_row(
            ["A320", data["AA320"][0], data["AA320"][1], data["AA320"][2], data["AA320"][3]])
        sheet3_T2.add_row(["", "", "", "", ""])
        sheet3_T2.add_row(
            ["Average Total", data["Average Total"][0], data["Average Total"][1], data["Average Total"][2], data["Average Total"][3]])

        # Print the table.
        print(sheet3_T2)

        # for key, value in data.items():
        #     print(key, value)

        # Variables to plot the chart.
        plotPie = []
        plotPieLabel = []

        # Append the data to plot the chart.
        for i in data:
            if i not in ["Total", "737", "AA320", "Average Total"]:
                plotPie.append(data[i][3])
                plotPieLabel.append(i)

        # Plot the chart
        plt.pie(np.array(plotPie), labels=plotPieLabel)
        # Show the chart.
        plt.show()