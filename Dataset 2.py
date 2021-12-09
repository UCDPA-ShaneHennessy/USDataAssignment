import xlrd
from prettytable import PrettyTable

import matplotlib.pyplot as plt
import numpy as np


class ShaneHennessyUSData:

    # This a a main method. this is the entry point of the program.
    if __name__ == "__main__":

        # Name of the file
        loc = "US Data Python Project.xlsx"

        # Open the excel file
        wb = xlrd.open_workbook(loc)

        # Variable to store the particular sheet. In this case, Sheet 2
        sheet = wb.sheet_by_index(1)

        # Header for the table
        sheet4Service = PrettyTable(["Type", "2017", "2018", "2019", "2020", "2021"])
        # Title for the table.
        sheet4Service.title = "Total in Service"

        # Header for another table
        sheet4Storage = PrettyTable(["Type", "2017", "2018", "2019", "2020", "2021"])
        # Title for another table.
        sheet4Storage.title = "Total in Storage"

        sheet4ServiceStorage = PrettyTable(["Type", "2017", "2018", "2019", "2020", "2021"])
        sheet4ServiceStorage.title = "Total in Service + Storage"

        sheet4FleetType = PrettyTable(["Type", "2017", "2018", "2019", "2020", "2021"])
        sheet4FleetType.title = "Proportion of Fleet Type in the Market"

        sheet4FleetType2 = PrettyTable(["Type", "2017", "2018", "2019", "2020", "2021"])
        sheet4FleetType2.title = "Proportion of Fleet Type in the Market"

        # Va,riables to store the data.
        dataService = {}
        dataStorage = {}
        dataServiceStorage = {}
        dataFleetType = {}

        # for i in range(1, sheet.nrows):
        #     print(sheet.cell(i, 0))
        #     pass

        # Loop to initialise the data variable as a list with default values.
        for i in range(1, sheet.nrows):
            valueType = str(sheet.cell_value(i, 1))

            # Check if the value is one in the list.
            if valueType in ["737 (CFMI)", "737 (JT8D)", "737 NG", "A319", "A320", "A321"]:
                # data[valueType + "Storage"] = 0
                # data[valueType + "InService"] = 0

                # Initialising the variables.
                dataService[valueType] = [0, 0, 0, 0, 0]
                dataStorage[valueType] = [0, 0, 0, 0, 0]
                dataServiceStorage[valueType] = [0, 0, 0, 0, 0]
                dataFleetType[valueType] = [0.0, 0.0, 0.0, 0.0, 0.0]

        # Variables to store totals.
        dataService["Total"] = [0, 0, 0, 0, 0]
        dataStorage["Total"] = [0, 0, 0, 0, 0]
        dataServiceStorage["Total"] = [0, 0, 0, 0, 0]
        dataFleetType["Total"] = [0.0, 0.0, 0.0, 0.0, 0.0]

        # Loop that does the calculation.
        for i in range(1, sheet.nrows):
            # Get the Type value
            valueType = str(sheet.cell_value(i, 1))
            # Get the status value
            status = str(sheet.cell_value(i, 3))

            if valueType in ["737 (CFMI)", "737 (JT8D)", "737 NG", "A319", "A320", "A321"]:

                # Check if the status value is Service of Storage.
                if str(status) == "Total In Service":
                    dataService[valueType][0] += int(sheet.cell_value(i, 4))
                    dataService[valueType][1] += int(sheet.cell_value(i, 5))
                    dataService[valueType][2] += int(sheet.cell_value(i, 6))
                    dataService[valueType][3] += int(sheet.cell_value(i, 7))
                    dataService[valueType][4] += int(sheet.cell_value(i, 8))
                elif str(status) == "Total In Storage":
                    dataStorage[valueType][0] += int(sheet.cell_value(i, 4))
                    dataStorage[valueType][1] += int(sheet.cell_value(i, 5))
                    dataStorage[valueType][2] += int(sheet.cell_value(i, 6))
                    dataStorage[valueType][3] += int(sheet.cell_value(i, 7))
                    dataStorage[valueType][4] += int(sheet.cell_value(i, 8))

        # Calculate the total for Storage
        for i in dataStorage:
            if i is not "Total":
                dataStorage["Total"][0] += dataStorage[i][0]
                dataStorage["Total"][1] += dataStorage[i][1]
                dataStorage["Total"][2] += dataStorage[i][2]
                dataStorage["Total"][3] += dataStorage[i][3]
                dataStorage["Total"][4] += dataStorage[i][4]

        # Calculate the total for Service
        for i in dataService:
            if i is not "Total":
                dataService["Total"][0] += dataService[i][0]
                dataService["Total"][1] += dataService[i][1]
                dataService["Total"][2] += dataService[i][2]
                dataService["Total"][3] += dataService[i][3]
                dataService["Total"][4] += dataService[i][4]

        # for key, value in dataService.items():
        #     print(key, value)
        #
        # print("\n\n")
        #
        # for key, value in dataStorage.items():
        #     print(key, value)

        # Add all the data to the table.
        sheet4Service.add_row(
            ["737 (CFMI)", dataService["737 (CFMI)"][0], dataService["737 (CFMI)"][1], dataService["737 (CFMI)"][2], dataService["737 (CFMI)"][3], dataService["737 (CFMI)"][4]])
        sheet4Service.add_row(
            ["737 (JT8D)", dataService["737 (JT8D)"][0], dataService["737 (JT8D)"][1], dataService["737 (JT8D)"][2], dataService["737 (JT8D)"][3], dataService["737 (JT8D)"][4]])
        sheet4Service.add_row(
            ["737 NG", dataService["737 NG"][0], dataService["737 NG"][1], dataService["737 NG"][2], dataService["737 NG"][3], dataService["737 NG"][4]])
        sheet4Service.add_row(
            ["A319", dataService["A319"][0], dataService["A319"][1], dataService["A319"][2], dataService["A319"][3], dataService["A319"][4]])
        sheet4Service.add_row(
            ["A320", dataService["A320"][0], dataService["A320"][1], dataService["A320"][2], dataService["A320"][3], dataService["A320"][4]])
        sheet4Service.add_row(
            ["A321", dataService["A321"][0], dataService["A321"][1], dataService["A321"][2], dataService["A321"][3], dataService["A321"][4]])
        sheet4Service.add_row(["", "", "", "", "", ""])
        sheet4Service.add_row(
            ["Total", dataService["Total"][0], dataService["Total"][1], dataService["Total"][2], dataService["Total"][3], dataService["Total"][4]])

        sheet4Storage.add_row(
            ["737 (CFMI)", dataStorage["737 (CFMI)"][0], dataStorage["737 (CFMI)"][1], dataStorage["737 (CFMI)"][2],
             dataStorage["737 (CFMI)"][3], dataStorage["737 (CFMI)"][4]])
        sheet4Storage.add_row(
            ["737 (JT8D)", dataStorage["737 (JT8D)"][0], dataStorage["737 (JT8D)"][1], dataStorage["737 (JT8D)"][2],
             dataStorage["737 (JT8D)"][3], dataStorage["737 (JT8D)"][4]])
        sheet4Storage.add_row(
            ["737 NG", dataStorage["737 NG"][0], dataStorage["737 NG"][1], dataStorage["737 NG"][2],
             dataStorage["737 NG"][3], dataStorage["737 NG"][4]])
        sheet4Storage.add_row(
            ["A319", dataStorage["A319"][0], dataStorage["A319"][1], dataStorage["A319"][2], dataStorage["A319"][3],
             dataStorage["A319"][4]])
        sheet4Storage.add_row(
            ["A320", dataStorage["A320"][0], dataStorage["A320"][1], dataStorage["A320"][2], dataStorage["A320"][3],
             dataStorage["A320"][4]])
        sheet4Storage.add_row(
            ["A321", dataStorage["A321"][0], dataStorage["A321"][1], dataStorage["A321"][2], dataStorage["A321"][3],
             dataStorage["A321"][4]])
        sheet4Storage.add_row(["", "", "", "", "", ""])
        sheet4Storage.add_row(
            ["Total", dataStorage["Total"][0], dataStorage["Total"][1], dataStorage["Total"][2],
             dataStorage["Total"][3], dataStorage["Total"][4]])

        # Loop to calculate both Service and Storage.
        for i in dataServiceStorage:
            dataServiceStorage[i][0] = dataService[i][0] + dataStorage[i][0]
            dataServiceStorage[i][1] = dataService[i][1] + dataStorage[i][1]
            dataServiceStorage[i][2] = dataService[i][2] + dataStorage[i][2]
            dataServiceStorage[i][3] = dataService[i][3] + dataStorage[i][3]
            dataServiceStorage[i][4] = dataService[i][4] + dataStorage[i][4]

        # Add the data to the table.
        sheet4ServiceStorage.add_row(
            ["737 (CFMI)", dataServiceStorage["737 (CFMI)"][0], dataServiceStorage["737 (CFMI)"][1], dataServiceStorage["737 (CFMI)"][2],
             dataServiceStorage["737 (CFMI)"][3], dataServiceStorage["737 (CFMI)"][4]])
        sheet4ServiceStorage.add_row(
            ["737 (JT8D)", dataServiceStorage["737 (JT8D)"][0], dataServiceStorage["737 (JT8D)"][1], dataServiceStorage["737 (JT8D)"][2],
             dataServiceStorage["737 (JT8D)"][3], dataServiceStorage["737 (JT8D)"][4]])
        sheet4ServiceStorage.add_row(
            ["737 NG", dataServiceStorage["737 NG"][0], dataServiceStorage["737 NG"][1], dataServiceStorage["737 NG"][2],
             dataServiceStorage["737 NG"][3], dataServiceStorage["737 NG"][4]])
        sheet4ServiceStorage.add_row(
            ["A319", dataServiceStorage["A319"][0], dataServiceStorage["A319"][1], dataServiceStorage["A319"][2], dataServiceStorage["A319"][3],
             dataServiceStorage["A319"][4]])
        sheet4ServiceStorage.add_row(
            ["A320", dataServiceStorage["A320"][0], dataServiceStorage["A320"][1], dataServiceStorage["A320"][2], dataServiceStorage["A320"][3],
             dataServiceStorage["A320"][4]])
        sheet4ServiceStorage.add_row(
            ["A321", dataServiceStorage["A321"][0], dataServiceStorage["A321"][1], dataServiceStorage["A321"][2], dataServiceStorage["A321"][3],
             dataServiceStorage["A321"][4]])
        sheet4ServiceStorage.add_row(["", "", "", "", "", ""])
        sheet4ServiceStorage.add_row(
            ["Total", dataServiceStorage["Total"][0], dataServiceStorage["Total"][1], dataServiceStorage["Total"][2],
             dataServiceStorage["Total"][3], dataServiceStorage["Total"][4]])

        # Calculate the fleet type.
        for i in dataFleetType:
            dataFleetType[i][0] = round(dataServiceStorage[i][0] / dataServiceStorage["Total"][0] * 100, 1)
            dataFleetType[i][1] = round(dataServiceStorage[i][1] / dataServiceStorage["Total"][1] * 100, 1)
            dataFleetType[i][2] = round(dataServiceStorage[i][2] / dataServiceStorage["Total"][2] * 100, 1)
            dataFleetType[i][3] = round(dataServiceStorage[i][3] / dataServiceStorage["Total"][3] * 100, 1)
            dataFleetType[i][4] = round(dataServiceStorage[i][4] / dataServiceStorage["Total"][4] * 100, 1)

        sheet4FleetType.add_row(
            ["737 (CFMI)", dataFleetType["737 (CFMI)"][0], dataFleetType["737 (CFMI)"][1],
             dataFleetType["737 (CFMI)"][2],
             dataFleetType["737 (CFMI)"][3], dataFleetType["737 (CFMI)"][4]])
        sheet4FleetType.add_row(
            ["737 (JT8D)", dataFleetType["737 (JT8D)"][0], dataFleetType["737 (JT8D)"][1],
             dataFleetType["737 (JT8D)"][2],
             dataFleetType["737 (JT8D)"][3], dataFleetType["737 (JT8D)"][4]])
        sheet4FleetType.add_row(
            ["737 NG", dataFleetType["737 NG"][0], dataFleetType["737 NG"][1],
             dataFleetType["737 NG"][2],
             dataFleetType["737 NG"][3], dataFleetType["737 NG"][4]])
        sheet4FleetType.add_row(
            ["A319", dataFleetType["A319"][0], dataFleetType["A319"][1], dataFleetType["A319"][2],
             dataFleetType["A319"][3],
             dataFleetType["A319"][4]])
        sheet4FleetType.add_row(
            ["A320", dataFleetType["A320"][0], dataFleetType["A320"][1], dataFleetType["A320"][2],
             dataFleetType["A320"][3],
             dataFleetType["A320"][4]])
        sheet4FleetType.add_row(
            ["A321", dataFleetType["A321"][0], dataFleetType["A321"][1], dataFleetType["A321"][2],
             dataFleetType["A321"][3],
             dataFleetType["A321"][4]])
        sheet4FleetType.add_row(["", "", "", "", "", ""])
        sheet4FleetType.add_row(
            ["Total", dataFleetType["Total"][0], dataFleetType["Total"][1], dataFleetType["Total"][2],
             dataFleetType["Total"][3], dataFleetType["Total"][4]])

        # Create and initialise other required variables.
        dataFleetType["737"] = [0.0, 0.0, 0.0, 0.0, 0.0]
        dataFleetType["AA320"] = [0.0, 0.0, 0.0, 0.0, 0.0]
        dataFleetType["Average Total"] = [0.0, 0.0, 0.0, 0.0, 0.0]

        # Loop to calculate total of 737 and A320.
        for i in dataFleetType:
            if i not in ["Total", "737", "AA320"]:
                if i[0] == "A":
                    dataFleetType["AA320"][0] += dataFleetType[i][0]
                    dataFleetType["AA320"][1] += dataFleetType[i][1]
                    dataFleetType["AA320"][2] += dataFleetType[i][2]
                    dataFleetType["AA320"][3] += dataFleetType[i][3]
                    dataFleetType["AA320"][4] += dataFleetType[i][4]
                else:
                    dataFleetType["737"][0] += dataFleetType[i][0]
                    dataFleetType["737"][1] += dataFleetType[i][1]
                    dataFleetType["737"][2] += dataFleetType[i][2]
                    dataFleetType["737"][3] += dataFleetType[i][3]
                    dataFleetType["737"][4] += dataFleetType[i][4]

        # Loop to calculate Average total of 737 and A320.
        for i in dataFleetType:
            if i in ["737", "AA320"]:
                dataFleetType["Average Total"][0] += dataFleetType[i][0]
                dataFleetType["Average Total"][1] += dataFleetType[i][1]
                dataFleetType["Average Total"][2] += dataFleetType[i][2]
                dataFleetType["Average Total"][3] += dataFleetType[i][3]
                dataFleetType["Average Total"][4] += dataFleetType[i][4]

        # Loop to round up the values.
        for i in dataFleetType:
            if i in ["737", "AA320", "Average Total"]:
                dataFleetType[i][0] = round(dataFleetType[i][0], 1)
                dataFleetType[i][1] = round(dataFleetType[i][1], 1)
                dataFleetType[i][2] = round(dataFleetType[i][2], 1)
                dataFleetType[i][3] = round(dataFleetType[i][3], 1)
                dataFleetType[i][4] = round(dataFleetType[i][4], 1)

        sheet4FleetType2.add_row(["737", dataFleetType["737"][0], dataFleetType["737"][1], dataFleetType["737"][2], dataFleetType["737"][3], dataFleetType["737"][4]])
        sheet4FleetType2.add_row(
            ["A320", dataFleetType["AA320"][0], dataFleetType["AA320"][1], dataFleetType["AA320"][2], dataFleetType["AA320"][3], dataFleetType["AA320"][4]])
        sheet4FleetType2.add_row(["", "", "", "", "", ""])
        sheet4FleetType2.add_row(
            ["Average Total", dataFleetType["Average Total"][0], dataFleetType["Average Total"][1], dataFleetType["Average Total"][2],
             dataFleetType["Average Total"][3], dataFleetType["Average Total"][4]])

        # Print all the tables
        print(sheet4Service)
        print(sheet4Storage)
        print(sheet4ServiceStorage)
        print(sheet4FleetType)
        print(sheet4FleetType2)

        # Variables to plot the graphs
        plotServiceStorage = {}
        plotFleetType = {}
        plotFleetType2 = {}

        # Loop to initialise the variables with empty list.
        for i in dataFleetType:
            if i not in ["Total", "737", "AA320", "Average Total"]:
                plotServiceStorage[i] = []
                plotFleetType[i] = []

            if i in ["737", "AA320"]:
                plotFleetType2[i] = []

        # Loop to store the data in the list.
        for i in dataFleetType:
            if i not in ["Total", "737", "AA320", "Average Total"]:
                plotServiceStorage[i].append(dataServiceStorage[i][0])
                plotServiceStorage[i].append(dataServiceStorage[i][1])
                plotServiceStorage[i].append(dataServiceStorage[i][2])
                plotServiceStorage[i].append(dataServiceStorage[i][3])
                plotServiceStorage[i].append(dataServiceStorage[i][4])

                plotFleetType[i].append(dataFleetType[i][0])
                plotFleetType[i].append(dataFleetType[i][1])
                plotFleetType[i].append(dataFleetType[i][2])
                plotFleetType[i].append(dataFleetType[i][3])
                plotFleetType[i].append(dataFleetType[i][4])

            if i in ["737", "AA320"]:
                plotFleetType2[i].append(dataFleetType[i][0])
                plotFleetType2[i].append(dataFleetType[i][1])
                plotFleetType2[i].append(dataFleetType[i][2])
                plotFleetType2[i].append(dataFleetType[i][3])
                plotFleetType2[i].append(dataFleetType[i][4])

        # for key, value in plotServiceStorage.items():
        #     print(key, value)
        #
        # print("\n\n")
        #
        # for key, value in plotFleetType.items():
        #     print(key, value)
        #
        # print("\n\n")
        #
        # for key, value in plotFleetType2.items():
        #     print(key, value)

        figure, axis = plt.subplots(2, 2)

        legend = []

        # Loop to plot the data.
        for i in plotServiceStorage:
            axis[0, 0].plot(np.array([2017, 2018, 2019, 2020, 2021]), np.array(plotServiceStorage[i]))
            # plt.plot(np.array([2017, 2018, 2019, 2020, 2021]), np.array(plotServiceStorage[i]))
            # legend.append(i)

        for i in plotFleetType:
            axis[0, 1].plot(np.array([2017, 2018, 2019, 2020, 2021]), np.array(plotFleetType[i]))

        for i in plotFleetType2:
            axis[1, 0].plot(np.array([2017, 2018, 2019, 2020, 2021]), np.array(plotFleetType2[i]))


        # Show the graph.
        plt.show()