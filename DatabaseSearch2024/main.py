# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import tkinter
import pandas as pd
import pyodbc
from tkinter import *
from tkinter import ttk


root = Tk()
# Defining default input frame
excel_frame = Frame(root)
# Defining input frame for manual access database paths
# access_frame = Frame(root)
error_string = StringVar()
root.title("GV database searcher V1.3")


def databaseSearch():
    # Obtain location of Excel file from input box:
    searchLocation = input_txt.get()
    # print(searchLocation)
    # Reset Error message string for each call of the Database search function.
    error_string.set("")
    # Read in the data from the Quote 2020 Form sheet of the inputted file.
    if searchLocation != "" and (searchLocation.endswith(".xlsx") or searchLocation.endswith(".xls")):
        # Confirm that file path was valid and that a connection to the Excel file was established. Otherwise, abort and return an error message.
        try:
            """
            # Code to confirm that Quote 2020 Form sheet exists in file.
            # Intended future usage is to enable program to also accept QUOTE 2023 FORM sheets.
            wb = load_workbook(searchLocation, read_only=True)
            if "QUOTE 2020 FORM" in wb.sheetnames:
                print("QUOTE 2020 FORM sheet exists in file")
            """
            try:
                # If inputted file path produces a valid Excel file with a sheet named "QUOTE 2020 FORM", read out its contents:
                df = pd.read_excel(searchLocation, sheet_name="QUOTE 2020 FORM")
                # Save the contents of the Parts Number and NSN columns in lists for reference.
                listPartsNumbers = df['PARTS NUMBER']
                listNSN = df['NSN']
                # Creating a list that will store the values of both lists as grouped tuples
                listCombo = []
                # Check if the length of the two lists are the same to avoid data mismatch or overflow errors:
                if len(listNSN) == len(listPartsNumbers):
                    # Iterate through the two lists
                    for i in range(len(listPartsNumbers)):
                        # Check and ensure that a value was entered in the Parts Number column or the NSN column:
                        if str(listPartsNumbers[i]) != 'nan' and str(listNSN[i]):
                            # Check and ensure a valid NSN was added:
                            if str(listNSN[i]) != 'nan' and str(listNSN[i]).startswith("PN") is not True and len(str(listNSN[i])) > 1:
                                # Save current items in partsNumbers and NSN in a tuple
                                tempTuple = (listPartsNumbers[i], listNSN[i])
                                # Append created tuple to the combined List item.
                                listCombo.append(tempTuple)
                            elif str(listNSN[i]).startswith(
                                    "PN") is True or len(str(listNSN[i])) == 1:  # If an invalid NSN of format PN[PartNumber] was placed in NSN cell
                                # Save current item in partsNumbers and invalid NSN notifier in a tuple
                                tempTuple = (listPartsNumbers[i], "Invalid NSN of format PN[PartNumber] entered")
                                # Append created tuple to the combined List item.
                                listCombo.append(tempTuple)
                            else:  # If no value was entered in the NSN cell
                                # Save current item in partsNumbers and empty NSN notifier in a tuple
                                tempTuple = (listPartsNumbers[i], "No NSN entered")
                                # Append created tuple to the combined List item.
                                listCombo.append(tempTuple)
                                # Define string to connect to on-device personal version of DB:
                connStr = (
                    r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};'
                    r'DBQ=C:\Users\mwright\Desktop\GV_MW_DB.accdb;'
                )
                try:
                    # Attempt connection to Database:
                    conn = pyodbc.connect(connStr)
                    # print("Connection Established to 2023 database")
                    cursor = conn.cursor()
                    search_results = []

                    if len(listCombo) != 0:
                        for i in range(len(listCombo)):
                            search_results_item = []
                            # Initialize variable to store string containing details of what part number and, if applicable, NSN number was searched for
                            entry_value = str(listCombo[i][0])
                            # print(entry_value)
                            if entry_value.find("(") != -1:
                                entry_value = entry_value.replace("(", " ")
                                entry_value = entry_value.replace(")", "")
                                # print(entry_value)
                            # If a valid NSN number was associated with the Part Number, update entry value variable to include the NSN number that was searched for:
                            if listCombo[i][1] != "No NSN entered" and listCombo[i][
                                1] != "Invalid NSN of format PN[PartNumber] entered":
                                entry_value = "(\'" + str(listCombo[i][0]) + "\', \'" + str(listCombo[i][1]) + "\')"
                                # print(entry_value)
                            # Define query to search 2DP HZ 2023 table:
                            query_TwoDPHZ23 = "SELECT DPRFQ, DATE, GVQ, NOTES, `Company Quotes` FROM `2DP HZ 2023` WHERE (`Parts Number`=\'" + str(
                                listCombo[i][0]) + "\'"
                            # Amend SQL query to include search for NSN numbers if applicable for given PN:
                            if listCombo[i][1] != "No NSN entered" and listCombo[i][1] != \
                                    "Invalid NSN of format PN[PartNumber] entered":
                                query_TwoDPHZ23 = query_TwoDPHZ23 + " OR `NSN`=\'" + str(
                                                listCombo[i][1]).rstrip() + "\'"
                            query_TwoDPHZ23 = query_TwoDPHZ23 + ") AND GVQ IS NOT NULL"
                            # print(query_TwoDPHZ23)
                            # Perform query to search 2DP HZ 2023 Storage table for matches:
                            df = pd.read_sql(query_TwoDPHZ23, conn)
                            # Format DATE column results to only show RFQ issue date (May need to revise this later-MW)
                            # print(df['DATE'].str.find("-"))
                            # print(df['DATE'].str[:7])
                            df['DATE'] = df['DATE'].str[:7]
                            # Format GVQ column results to read as currency values of form '$x.00'
                            df['GVQ'] = df['GVQ'].apply(lambda x: "${:.2f}".format((x)))
                            # Store query results as a list of values
                            values_list_TwoDPHZ2023 = df.values.tolist()
                            # print(values_list_TwoDPHZ2023)
                            # Confirm that matching values were found in the table:
                            if len(values_list_TwoDPHZ2023) != 0:
                                search_results_item.append(values_list_TwoDPHZ2023)

                            # Define query to search 2DP HZ 2023 table for null price matches:
                            query_TwoDPHZ23_Null = "SELECT DPRFQ, DATE FROM `2DP HZ 2023` WHERE (`Parts Number`=\'" + str(
                                listCombo[i][0]) + "\'"
                            # Amend SQL query to include search for NSN numbers if applicable for given PN:
                            if listCombo[i][1] != "No NSN entered" and listCombo[i][1] \
                                    != "Invalid NSN of format PN[PartNumber] entered":
                                query_TwoDPHZ23_Null = query_TwoDPHZ23_Null + " OR `NSN`=\'" + str(
                                    listCombo[i][1]).rstrip() + "\'"
                            query_TwoDPHZ23_Null = query_TwoDPHZ23_Null + ") AND GVQ IS NULL"
                            # print(query_TwoDPHZ23_Null)
                            # Perform query to search 2DP HZ 2023 Storage table for matches:
                            """
                            df = pd.read_sql(query_TwoDPHZ23_Null, conn)
                            # Format DATE column results to only show RFQ issue date (May need to revise this later-MW)
                            # print(df['DATE'].str.find("-"))
                            # print(df['DATE'].str[:7])
                            df['DATE'] = df['DATE'].str[:7]
                            # Store query results as a list of values
                            values_list_TwoDPHZ2023_Null = df.values.tolist()
                            # print(values_list_TwoDPHZ2023)
                            # Confirm that matching values were found in the table:
                            if len(values_list_TwoDPHZ2023_Null) != 0:
                                search_results_item.append(values_list_TwoDPHZ2023_Null)
                            """
                            result = cursor.execute(query_TwoDPHZ23_Null)
                            cursor_2DPHZ23_Null_RList = []
                            for row in result:
                                # print(row.DPRFQ)
                                dprfq_Mod = row.DPRFQ + " (2023)"
                                # print(dprfq_Mod)
                                if cursor_2DPHZ23_Null_RList.count(dprfq_Mod + ", No bid") == 0:
                                    cursor_2DPHZ23_Null_RList.append(dprfq_Mod + ", No bid")
                            if len(cursor_2DPHZ23_Null_RList) != 0:
                                # print(cursor_2DPHZ23_Null_RList)
                                search_results_item.append(cursor_2DPHZ23_Null_RList)

                            # Define query to search 2DP MAX 2023 table:
                            query_TwoDPMAX23 = "SELECT DPRFQ, DATE, ALMCQ, SUGGEST, NOTES, `Company Quotes` FROM `2DP MAX 2023` WHERE (`Parts Number`=\'" + str(
                                listCombo[i][0]) + "\'"
                            # Amend SQL query to include search for NSN numbers if applicable for given PN:
                            if listCombo[i][1] != "No NSN entered" and listCombo[i][1] \
                                    != "Invalid NSN of format PN[PartNumber] entered":
                                query_TwoDPMAX23 = query_TwoDPMAX23 + " OR `NSN`=\'" + str(
                                    listCombo[i][1]).rstrip() + "\'"
                            query_TwoDPMAX23 = query_TwoDPMAX23 + ") AND ALMCQ IS NOT NULL"
                            # print(query_TwoDPMAX23)
                            # Perform query to search 2DP HZ 2023 Storage table for matches:
                            df = pd.read_sql(query_TwoDPMAX23, conn)
                            # Format DATE column results to only show RFQ issue date (May need to revise this later-MW)
                            # print(df['DATE'].str.find("-"))
                            # print(df['DATE'].str[:7])
                            df['DATE'] = df['DATE'].str[:7]
                            # Format ALMCQ and SUGGEST column results to read as currency values of form '$x.00'
                            df['ALMCQ'] = df['ALMCQ'].apply(lambda x: "${:.2f}".format((x)))
                            df['SUGGEST'] = df['SUGGEST'].apply(lambda x: "${:.2f}".format((x)))
                            # Store query results as a list of values
                            values_list_TwoDPMAX2023 = df.values.tolist()
                            # print(values_list_TwoDPMAX2023)
                            # Confirm that matching values were found in the table:
                            if len(values_list_TwoDPMAX2023) != 0:
                                search_results_item.append(values_list_TwoDPMAX2023)

                            # Define query to search 2DP MAX 2023 table for null price matches:
                            query_TwoDPMAX23_Null = "SELECT DPRFQ, DATE FROM `2DP MAX 2023` WHERE (`Parts Number`=\'" + str(
                                listCombo[i][0]) + "\'"
                            # Amend SQL query to include search for NSN numbers if applicable for given PN:
                            if listCombo[i][1] != "No NSN entered" and listCombo[i][1] \
                                    != "Invalid NSN of format PN[PartNumber] entered":
                                query_TwoDPMAX23_Null = query_TwoDPMAX23_Null + " OR `NSN`=\'" + str(
                                    listCombo[i][1]).rstrip() + "\'"
                            query_TwoDPMAX23_Null = query_TwoDPMAX23_Null + ") AND ALMCQ IS NULL"
                            # print(query_TwoDPMAX23_Null)
                            # Perform query to search 2DP HZ 2023 Storage table for matches:
                            """
                            df = pd.read_sql(query_TwoDPMAX23_Null, conn)
                            # Format DATE column results to only show RFQ issue date (May need to revise this later-MW)
                            # print(df['DATE'].str.find("-"))
                            # print(df['DATE'].str[:7])
                            df['DATE'] = df['DATE'].str[:7]
                            # Store query results as a list of values
                            values_list_TwoDPMAX2023_Null = df.values.tolist()
                            # print(values_list_TwoDPMAX2023)
                            # Confirm that matching values were found in the table:
                            if len(values_list_TwoDPMAX2023_Null) != 0:
                                search_results_item.append(values_list_TwoDPMAX2023_Null)
                            """
                            result = cursor.execute(query_TwoDPMAX23_Null)
                            cursor_2DPMax23_Null_RList = []
                            for row in result:
                                # print(row.DPRFQ)
                                dprfq_Mod = row.DPRFQ + " (2023)"
                                # print(dprfq_Mod)
                                if cursor_2DPMax23_Null_RList.count(dprfq_Mod + ", No bid") == 0:
                                    cursor_2DPMax23_Null_RList.append(dprfq_Mod + ", No bid")
                            if len(cursor_2DPMax23_Null_RList) != 0:
                                # print(cursor_2DPMax23_Null_RList)
                                search_results_item.append(cursor_2DPMax23_Null_RList)

                            # Define query to search 2AD Max 2023 table:
                            query_TwoADMAX23 = "SELECT DPRFQ, DATE, ALMCQ, SUGGEST, NOTES, `Company Quotes` FROM `2AD Max 2023` WHERE (`Parts Number`=\'" + str(
                                listCombo[i][0]) + "\'"
                            # Amend SQL query to include search for NSN numbers if applicable for given PN:
                            if listCombo[i][1] != "No NSN entered" and listCombo[i][1] \
                                    != "Invalid NSN of format PN[PartNumber] entered":
                                query_TwoADMAX23 = query_TwoADMAX23 + " OR `NSN`=\'" + str(
                                    listCombo[i][1]).rstrip() + "\'"
                            query_TwoADMAX23 = query_TwoADMAX23 + ") AND ALMCQ IS NOT NULL"
                            # print(query_TwoADMAX23)
                            # Perform query to search 2AD Max 2023 Storage table for matches:
                            df = pd.read_sql(query_TwoADMAX23, conn)
                            # Format DATE column results to only show RFQ issue date (May need to revise this later-MW)
                            # print(df['DATE'].str.find("-"))
                            # print(df['DATE'].str[:7])
                            df['DATE'] = df['DATE'].str[:7]
                            # Format ALMCQ and SUGGEST column results to read as currency values of form '$x.00'
                            df['ALMCQ'] = df['ALMCQ'].apply(lambda x: "${:.2f}".format((x)))
                            df['SUGGEST'] = df['SUGGEST'].apply(lambda x: "${:.2f}".format((x)))
                            # Store query results as a list of values
                            values_list_TwoADMax2023 = df.values.tolist()
                            # print(values_list_TwoADMax2023)
                            # Confirm that matching values were found in the table:
                            if len(values_list_TwoADMax2023) != 0:
                                # search_resultsTwoADMaxB = ["2AD Max 2023", values_list_TwoADMax2023]
                                search_results_item.append(values_list_TwoADMax2023)

                            # Define query to search 2AD Max 2023 table for null price matches:
                            query_TwoADMAX23_Null = "SELECT DPRFQ, DATE FROM `2AD Max 2023` WHERE (`Parts Number`=\'" + str(
                                listCombo[i][0]) + "\'"
                            # Amend SQL query to include search for NSN numbers if applicable for given PN:
                            if listCombo[i][1] != "No NSN entered" and listCombo[i][1] \
                                    != "Invalid NSN of format PN[PartNumber] entered":
                                query_TwoADMAX23_Null = query_TwoADMAX23_Null + " OR `NSN`=\'" + str(
                                    listCombo[i][1]).rstrip() + "\'"
                            query_TwoADMAX23_Null = query_TwoADMAX23_Null + ") AND ALMCQ IS NULL"
                            # print(query_TwoADMAX23_Null)
                            # Perform query to search 2AD Max 2023 Storage table for matches:
                            """
                            df = pd.read_sql(query_TwoADMAX23_Null, conn)
                            # Format DATE column results to only show RFQ issue date (May need to revise this later-MW)
                            # print(df['DATE'].str.find("-"))
                            # print(df['DATE'].str[:7])
                            df['DATE'] = df['DATE'].str[:7]
                            # Store query results as a list of values
                            values_list_TwoADMax2023_Null = df.values.tolist()
                            # print(values_list_TwoADMax2023)
                            # Confirm that matching values were found in the table:
                            if len(values_list_TwoADMax2023_Null) != 0:
                                search_results_item.append(values_list_TwoADMax2023_Null)
                            """
                            result = cursor.execute(query_TwoADMAX23_Null)
                            cursor_2ADMax23_Null_RList = []
                            for row in result:
                                # print(row.DPRFQ)
                                dprfq_Mod = row.DPRFQ + " (2023)"
                                # print(dprfq_Mod)
                                if cursor_2ADMax23_Null_RList.count(dprfq_Mod + ", No bid") == 0:
                                    cursor_2ADMax23_Null_RList.append(dprfq_Mod + ", No bid")
                            if len(cursor_2ADMax23_Null_RList) != 0:
                                # print(cursor_2ADMax23_Null_RList)
                                search_results_item.append(cursor_2ADMax23_Null_RList)

                            # Define query to search 3DP Max 2023 table:
                            query_ThreeDPmax23 = "SELECT DPRFQ, DATE, ALMCQ, SUGGEST, NOTES, `Company Quotes` FROM `3DP Max 2023` WHERE (`Parts Number`=\'" + str(
                                listCombo[i][0]) + "\'"
                            # Amend SQL query to include search for NSN numbers if applicable for given PN:
                            if listCombo[i][1] != "No NSN entered" and listCombo[i][
                                1] != "Invalid NSN of format PN[PartNumber] entered":
                                query_ThreeDPmax23 = query_ThreeDPmax23 + " OR `NSN`=\'" + str(
                                    listCombo[i][1]).rstrip() + "\'"
                            query_ThreeDPmax23 = query_ThreeDPmax23 + ") AND ALMCQ IS NOT NULL"
                            # Perform query to search 3DP Max 2023 table for matches:
                            # print(query_ThreeDPmax23)
                            # print(query_ThreeDPmax23)
                            df = pd.read_sql(query_ThreeDPmax23, conn)
                            # Format DATE column results to only show RFQ issue date (May need to revise this later-MW)
                            # print(df['DATE'].str.find("-"))
                            # print(df['DATE'].str[:7])
                            df['DATE'] = df['DATE'].str[:7]
                            # Format ALMCQ and SUGGEST column results to read as currency values of form '$x.00'
                            df['ALMCQ'] = df['ALMCQ'].apply(lambda x: "${:.2f}".format((x)))
                            df['SUGGEST'] = df['SUGGEST'].apply(lambda x: "${:.2f}".format((x)))
                            # Store query results as a list of values
                            values_list_ThreeDPmax23 = df.values.tolist()
                            if len(values_list_ThreeDPmax23) != 0:
                                # If any rows in 3AD Max Temp Moving Storage table contained any matching values,
                                # return the contents of these rows to the user.
                                # search_results_ThreeDPmax23 = ["3DP 2023", values_list_ThreeDPmax23]
                                search_results_item.append(values_list_ThreeDPmax23)

                            # Define query to search 3DP Max 2023 table for null price matches:
                            query_ThreeDPmax23_Null = "SELECT DPRFQ, DATE FROM `3DP Max 2023` WHERE (`Parts Number`=\'" + str(
                                listCombo[i][0]) + "\'"
                            # Amend SQL query to include search for NSN numbers if applicable for given PN:
                            if listCombo[i][1] != "No NSN entered" and listCombo[i][
                                1] != "Invalid NSN of format PN[PartNumber] entered":
                                query_ThreeDPmax23_Null = query_ThreeDPmax23_Null + " OR `NSN`=\'" + str(
                                    listCombo[i][1]).rstrip() + "\'"
                            query_ThreeDPmax23_Null = query_ThreeDPmax23_Null + ") AND ALMCQ IS NULL"
                            # Perform query to search 3DP Max 2023 table for matches:
                            # print(query_ThreeDPmax23_Null)
                            """
                            df = pd.read_sql(query_ThreeDPmax23_Null, conn)
                            # Format DATE column results to only show RFQ issue date (May need to revise this later-MW)
                            # print(df['DATE'].str.find("-"))
                            # print(df['DATE'].str[:7])
                            df['DATE'] = df['DATE'].str[:7]
                            # Store query results as a list of values
                            values_list_ThreeDPmax23_Null = df.values.tolist()
                            if len(values_list_ThreeDPmax23_Null) != 0:
                                # If any rows in 3AD Max Temp Moving Storage table contained any matching values,
                                # return the contents of these rows to the user.
                                # search_results_ThreeDPmax23 = ["3DP 2023", values_list_ThreeDPmax23]
                                search_results_item.append(values_list_ThreeDPmax23_Null)
                            """
                            result = cursor.execute(query_ThreeDPmax23_Null)
                            cursor_3DPMax23_Null_RList = []
                            for row in result:
                                # print(row.DPRFQ)
                                dprfq_Mod = row.DPRFQ + " (2023)"
                                # print(dprfq_Mod)
                                if cursor_3DPMax23_Null_RList.count(dprfq_Mod + ", No bid") == 0:
                                    cursor_3DPMax23_Null_RList.append(dprfq_Mod + ", No bid")
                            if len(cursor_3DPMax23_Null_RList) != 0:
                                # print(cursor_3DPMax23_Null_RList)
                                search_results_item.append(cursor_3DPMax23_Null_RList)

                            # Define query to search 4DP Max 2023 table:
                            query_FourDPmax23 = "SELECT DPRFQ, DATE, ALMCQ, SUGGEST, NOTES, `Company Quotes` FROM `4DP Max 2023` WHERE (`Parts Number`=\'" + str(
                                listCombo[i][0]) + "\'"
                            # Amend SQL query to include search for NSN numbers if applicable for given PN:
                            if listCombo[i][1] != "No NSN entered" and listCombo[i][
                                1] != "Invalid NSN of format PN[PartNumber] entered":
                                query_FourDPmax23 = query_FourDPmax23 + " OR `NSN`=\'" + str(
                                    listCombo[i][1]).rstrip() + "\'"
                            query_FourDPmax23 = query_FourDPmax23 + ") AND ALMCQ IS NOT NULL"
                            # print(query_FourDPmax23)
                            # Perform query to search 4DP Max 2023 table for matches:
                            # print(query_FourDPmax23)
                            df = pd.read_sql(query_FourDPmax23, conn)
                            # Format DATE column results to only show RFQ issue date (May need to revise this later-MW)
                            # print(df['DATE'].str.find("-"))
                            # print(df['DATE'].str[:7])
                            df['DATE'] = df['DATE'].str[:7]
                            # Format ALMCQ and SUGGEST column results to read as currency values of form '$x.00'
                            df['ALMCQ'] = df['ALMCQ'].apply(lambda x: "${:.2f}".format((x)))
                            df['SUGGEST'] = df['SUGGEST'].apply(lambda x: "${:.2f}".format((x)))
                            # Store query results as a list of values
                            values_list_FourDPmax23 = df.values.tolist()
                            if len(values_list_FourDPmax23) != 0:
                                # If any rows in 4DP Max 2023 table contained any matching values,
                                # return the contents of these rows to the user.
                                # search_results_ThreeDPmax23 = ["3DP 2023", values_list_ThreeDPmax23]
                                search_results_item.append(values_list_FourDPmax23)

                            # Define query to search 4DP Max 2023 table for Null price matches
                            query_FourDPmax23_Null = "SELECT DPRFQ, DATE FROM `4DP Max 2023` WHERE (`Parts Number`=\'" + str(
                                listCombo[i][0]) + "\'"
                            # Amend SQL query to include search for NSN numbers if applicable for given PN:
                            if listCombo[i][1] != "No NSN entered" and listCombo[i][
                                1] != "Invalid NSN of format PN[PartNumber] entered":
                                query_FourDPmax23_Null = query_FourDPmax23_Null + " OR `NSN`=\'" + str(
                                    listCombo[i][1]).rstrip() + "\'"
                            query_FourDPmax23_Null = query_FourDPmax23_Null + ") AND ALMCQ IS NULL"
                            # print(query_FourDPmax23_Null)
                            # Perform query to search 4DP Max 2023 table for matches:
                            # print(query_FourDPmax23)
                            """
                            df = pd.read_sql(query_FourDPmax23_Null, conn)
                            # Format DATE column results to only show RFQ issue date (May need to revise this later-MW)
                            # print(df['DATE'].str.find("-"))
                            # print(df['DATE'].str[:7])
                            df['DATE'] = df['DATE'].str[:7]
                            values_list_FourDPmax23_Null = df.values.tolist()
                            if len(values_list_FourDPmax23_Null) != 0:
                                search_results_item.append(values_list_FourDPmax23_Null)
                            """
                            result = cursor.execute(query_FourDPmax23_Null)
                            cursor_4DPMax23_Null_RList = []
                            for row in result:
                                # print(row.DPRFQ)
                                dprfq_Mod = row.DPRFQ + " (2023)"
                                # print(dprfq_Mod)
                                if cursor_4DPMax23_Null_RList.count(dprfq_Mod + ", No bid") == 0:
                                    cursor_4DPMax23_Null_RList.append(dprfq_Mod + ", No bid")
                            if len(cursor_4DPMax23_Null_RList) != 0:
                                # print(cursor_4DPMax23_Null_RList)
                                search_results_item.append(cursor_4DPMax23_Null_RList)

                            # Define query to search 3AD Max 2023 table:
                            query_ThreeADMAX23 = "SELECT DPRFQ, DATE, ALMCQ, SUGGEST, NOTES, `Company Quotes` FROM `3AD Max Storage 2023` WHERE (`Parts Number`=\'" + str(
                                listCombo[i][0]) + "\'"
                            # Amend SQL query to include search for NSN numbers if applicable for given PN:
                            if listCombo[i][1] != "No NSN entered" and listCombo[i][1] \
                                    != "Invalid NSN of format PN[PartNumber] entered":
                                query_ThreeADMAX23 = query_ThreeADMAX23 + " OR `NSN`=\'" + str(
                                    listCombo[i][1]).rstrip() + "\'"
                            query_ThreeADMAX23 = query_ThreeADMAX23 + ") AND ALMCQ IS NOT NULL"
                            # print(query_ThreeADMAX23)
                            # Perform query to search 3AD Max 2023 Storage table for matches:
                            df = pd.read_sql(query_ThreeADMAX23, conn)
                            # Format DATE column results to only show RFQ issue date (May need to revise this later-MW)
                            # print(df['DATE'].str.find("-"))
                            # print(df['DATE'].str[:7])
                            df['DATE'] = df['DATE'].str[:7]
                            # Format ALMCQ and SUGGEST column results to read as currency values of form '$x.00'
                            df['ALMCQ'] = df['ALMCQ'].apply(lambda x: "${:.2f}".format((x)))
                            df['SUGGEST'] = df['SUGGEST'].apply(lambda x: "${:.2f}".format((x)))
                            # Store query results as a list of values
                            values_list_ThreeADmax2023 = df.values.tolist()
                            # print(values_list_ThreeADmax2023)
                            # Confirm that matching values were found in the table:
                            if len(values_list_ThreeADmax2023) != 0:
                                # search_resultsThreeADMaxB = ["3AD Max 2023", values_list_ThreeADmax2023]
                                search_results_item.append(values_list_ThreeADmax2023)

                            # Define query to search 3AD Max 2023 table for null price matches:
                            query_ThreeADMAX23_Null = "SELECT DPRFQ, DATE FROM `3AD Max Storage 2023` WHERE (`Parts Number`=\'" + str(
                                listCombo[i][0]) + "\'"
                            # Amend SQL query to include search for NSN numbers if applicable for given PN:
                            if listCombo[i][1] != "No NSN entered" and listCombo[i][1] \
                                    != "Invalid NSN of format PN[PartNumber] entered":
                                query_ThreeADMAX23_Null = query_ThreeADMAX23_Null + " OR `NSN`=\'" + str(
                                    listCombo[i][1]).rstrip() + "\'"
                            query_ThreeADMAX23_Null = query_ThreeADMAX23_Null + ") AND ALMCQ IS NULL"
                            # print(query_ThreeADMAX23_Null)
                            """
                            # Perform query to search 3AD Max 2023 Storage table for matches:
                            df = pd.read_sql(query_ThreeADMAX23_Null, conn)
                            # Format DATE column results to only show RFQ issue date (May need to revise this later-MW)
                            # print(df['DATE'].str.find("-"))
                            # print(df['DATE'].str[:7])
                            df['DATE'] = df['DATE'].str[:7]
                            # Store query results as a list of values
                            values_list_ThreeADmax2023_Null = df.values.tolist()
                            # print(values_list_ThreeADmax2023_Null)
                            # Confirm that matching values were found in the table:
                            if len(values_list_ThreeADmax2023_Null) != 0:
                                # search_resultsThreeADMaxB = ["3AD Max 2023", values_list_ThreeADmax2023_Null]
                                search_results_item.append(values_list_ThreeADmax2023_Null)
                            """
                            # Experimental code to condense and reformat row strings:
                            result = cursor.execute(query_ThreeADMAX23_Null)
                            cursor_3ADMax23_Null_RList = []
                            for row in result:
                                # print(row.DPRFQ)
                                dprfq_Mod = row.DPRFQ + " (2023)";
                                # print(dprfq_Mod)
                                if cursor_3ADMax23_Null_RList.count(dprfq_Mod + ", No bid") == 0:
                                    cursor_3ADMax23_Null_RList.append(dprfq_Mod + ", No bid")
                            if len(cursor_3ADMax23_Null_RList) != 0:
                                # print(cursor_3ADMax23_Null_RList)
                                search_results_item.append(cursor_3ADMax23_Null_RList)

                            # Define query to search 4AD Max 2023 table:
                            query_FourADMAX23 = "SELECT DPRFQ, DATE, ALMCQ, SUGGEST, NOTES, `Company Quotes` FROM `4AD Max 2023 Storage` WHERE (`Parts Number`=\'" + str(
                                listCombo[i][0]) + "\'"
                            # Amend SQL query to include search for NSN numbers if applicable for given PN:
                            if listCombo[i][1] != "No NSN entered" and listCombo[i][1] \
                                    != "Invalid NSN of format PN[PartNumber] entered":
                                query_FourADMAX23 = query_FourADMAX23 + " OR `NSN`=\'" + str(
                                    listCombo[i][1]).rstrip() + "\'"
                            query_FourADMAX23 = query_FourADMAX23 + ") AND ALMCQ IS NOT NULL"
                            # print(query_FourADMAX23)
                            # Perform query to search 4AD Max 2023 Storage table for matches:
                            df = pd.read_sql(query_FourADMAX23, conn)
                            # Format DATE column results to only show RFQ issue date (May need to revise this later-MW)
                            # print(df['DATE'].str.find("-"))
                            # print(df['DATE'].str[:7])
                            df['DATE'] = df['DATE'].str[:7]
                            # Format ALMCQ and SUGGEST column results to read as currency values of form '$x.00'
                            df['ALMCQ'] = df['ALMCQ'].apply(lambda x: "${:.2f}".format((x)))
                            df['SUGGEST'] = df['SUGGEST'].apply(lambda x: "${:.2f}".format((x)))
                            # Store query results as a list of values
                            values_list_FourADmax2023 = df.values.tolist()
                            # print(values_list_FourADmax2023)
                            # Confirm that matching values were found in the table:
                            if len(values_list_FourADmax2023) != 0:
                                # search_resultsFourADMaxB = ["4AD Max 2023", values_list_FourADmax2023]
                                search_results_item.append(values_list_FourADmax2023)

                            # Define query to search 4AD Max 2023 table for null price matches:
                            query_FourADMAX23_Null = "SELECT DPRFQ, DATE FROM `4AD Max 2023 Storage` WHERE (`Parts Number`=\'" + str(
                                listCombo[i][0]) + "\'"
                            # Amend SQL query to include search for NSN numbers if applicable for given PN:
                            if listCombo[i][1] != "No NSN entered" and listCombo[i][1] \
                                    != "Invalid NSN of format PN[PartNumber] entered":
                                query_FourADMAX23_Null = query_FourADMAX23_Null + " OR `NSN`=\'" + str(
                                    listCombo[i][1]).rstrip() + "\'"
                            query_FourADMAX23_Null = query_FourADMAX23_Null + ") AND ALMCQ IS NULL"
                            # print(query_FourADMAX23_Null)
                            """
                            # Perform query to search 4AD Max 2023 Storage table for matches:
                            df = pd.read_sql(query_FourADMAX23_Null, conn)
                            # Format DATE column results to only show RFQ issue date (May need to revise this later-MW)
                            # print(df['DATE'].str.find("-"))
                            # print(df['DATE'].str[:7])
                            df['DATE'] = df['DATE'].str[:7]
                            # Store query results as a list of values
                            values_list_FourADmax2023_Null = df.values.tolist()
                            # print(values_list_FourADmax2023)
                            # Confirm that matching values were found in the table:
                            if len(values_list_FourADmax2023_Null) != 0:
                                # search_resultsFourADMaxB = ["4AD Max 2023", values_list_FourADmax2023]
                                search_results_item.append(values_list_FourADmax2023_Null)
                            """
                            # Experimental code to condense and reformat row strings:
                            result = cursor.execute(query_FourADMAX23_Null)
                            cursor_4ADMax23_Null_RList = []
                            for row in result:
                                # print(row.DPRFQ)
                                dprfq_Mod = row.DPRFQ + "(2023)"
                                # print(dprfq_Mod)
                                if cursor_4ADMax23_Null_RList.count(dprfq_Mod + ", No bid") == 0:
                                    cursor_4ADMax23_Null_RList.append(dprfq_Mod + ", No bid")
                            if len(cursor_4ADMax23_Null_RList) != 0:
                                # print(cursor_4ADMax23_Null_RList)
                                search_results_item.append(cursor_4ADMax23_Null_RList)

                            # Define query to search Kuji AD 2023 table:
                            query_KujiAD23 = "SELECT DPRFQ, DATE, ALMCQ, SUGGEST, NOTES, `Company Quotes` FROM `Kuji AD 2023` WHERE (`Parts Number`=\'" + str(
                                listCombo[i][0]) + "\'"
                            # Amend SQL query to include search for NSN numbers if applicable for given PN:
                            if listCombo[i][1] != "No NSN entered" and listCombo[i][1] \
                                    != "Invalid NSN of format PN[PartNumber] entered":
                                query_KujiAD23 = query_KujiAD23 + " OR `NSN`=\'" + str(
                                    listCombo[i][1]).rstrip() + "\'"
                            # Perform query to search 4AD Max 2023 Storage table for matches:
                            query_KujiAD23 = query_KujiAD23 + ") AND ALMCQ IS NOT NULL"
                            # print(query_KujiAD23)
                            df = pd.read_sql(query_KujiAD23, conn)
                            # Format DATE column results to only show RFQ issue date (May need to revise this later-MW)
                            # print(df['DATE'].str.find("-"))
                            # print(df['DATE'].str[:7])
                            df['DATE'] = df['DATE'].str[:7]
                            # Format ALMCQ and SUGGEST column results to read as currency values of form '$x.00'
                            df['ALMCQ'] = df['ALMCQ'].apply(lambda x: "${:.2f}".format((x)))
                            df['SUGGEST'] = df['SUGGEST'].apply(lambda x: "${:.2f}".format((x)))
                            # Store query results as a list of values
                            values_list_KujiAD2023 = df.values.tolist()
                            # print(values_list_KujiAD2023)
                            # Confirm that matching values were found in the table:
                            if len(values_list_KujiAD2023) != 0:
                                # search_resultsKujiAD23 = ["Kuji AD 2023", values_list_KujiAD2023]
                                search_results_item.append(values_list_KujiAD2023)

                            # Define query to search Kuji AD 2023 table for Null price matches:
                            query_KujiAD23_Null = "SELECT DPRFQ, DATE FROM `Kuji AD 2023` WHERE (`Parts Number`=\'" + str(
                                listCombo[i][0]) + "\'"
                            # Amend SQL query to include search for NSN numbers if applicable for given PN:
                            if listCombo[i][1] != "No NSN entered" and listCombo[i][1] \
                                    != "Invalid NSN of format PN[PartNumber] entered":
                                query_KujiAD23_Null = query_KujiAD23_Null + " OR `NSN`=\'" + str(
                                    listCombo[i][1]).rstrip() + "\'"
                            # Perform query to search 4AD Max 2023 Storage table for matches:
                            query_KujiAD23_Null = query_KujiAD23_Null + ") AND ALMCQ IS NULL"
                            # print(query_KujiAD23_Null)
                            """
                            # Old obselete code for searching for and returning matches:
                            df = pd.read_sql(query_KujiAD23_Null, conn)
                            # Format DATE column results to only show RFQ issue date (May need to revise this later-MW)
                            # print(df['DATE'].str.find("-"))
                            # print(df['DATE'].str[:7])
                            df['DATE'] = df['DATE'].str[:7]
                            # Store query results as a list of values
                            values_list_KujiAD2023_Null = df.values.tolist()
                            # print(values_list_KujiAD2023_Null)
                            # Confirm that matching values were found in the table:
                            if len(values_list_KujiAD2023_Null) != 0:
                                # search_resultsKujiAD23_Null = ["Kuji AD 2023", values_list_KujiAD2023_Null]
                                search_results_item.append(values_list_KujiAD2023_Null)
                            """
                            # Experimental code to condense and reformat row strings:
                            result = cursor.execute(query_KujiAD23_Null)
                            cursor_KujiAD23_Null_RList = []
                            for row in result:
                                # print(row.DPRFQ)
                                dprfq_Mod = row.DPRFQ + " (2023)"
                                # print(dprfq_Mod)
                                if cursor_KujiAD23_Null_RList.count(dprfq_Mod + ", No bid") == 0:
                                    cursor_KujiAD23_Null_RList.append(dprfq_Mod + ", No bid")
                            if len(cursor_KujiAD23_Null_RList) != 0:
                                # print(cursor_KujiAD23_Null_RList)
                                search_results_item.append(cursor_KujiAD23_Null_RList)

                            # Define query to search Kuji R 2023 table:
                            query_KujiR23 = "SELECT DPRFQ, DATE, ALMCQ, SUGGEST, NOTES, `Company Quotes` FROM `Kuji R Storage 2023` WHERE (`Parts Number`=\'" + str(
                                listCombo[i][0]) + "\'"
                            # Amend SQL query to include search for NSN numbers if applicable for given PN:
                            if listCombo[i][1] != "No NSN entered" and listCombo[i][1] \
                                    != "Invalid NSN of format PN[PartNumber] entered":
                                query_KujiR23 = query_KujiR23 + " OR `NSN`=\'" + str(
                                    listCombo[i][1]).rstrip() + "\'"
                            # Perform query to search 4AD Max 2023 Storage table for matches:
                            query_KujiR23 = query_KujiR23 + ") AND ALMCQ IS NOT NULL"
                            # print(query_KujiR23)
                            df = pd.read_sql(query_KujiR23, conn)
                            # Format DATE column results to only show RFQ issue date (May need to revise this later-MW)
                            # print(df['DATE'].str.find("-"))
                            # print(df['DATE'].str[:7])
                            df['DATE'] = df['DATE'].str[:7]
                            # Format ALMCQ and SUGGEST column results to read as currency values of form '$x.00'
                            df['ALMCQ'] = df['ALMCQ'].apply(lambda x: "${:.2f}".format((x)))
                            df['SUGGEST'] = df['SUGGEST'].apply(lambda x: "${:.2f}".format((x)))
                            # Store query results as a list of values
                            values_list_KujiR2023 = df.values.tolist()
                            # print(values_list_KujiR2023)
                            # Confirm that matching values were found in the table:
                            if len(values_list_KujiR2023) != 0:
                                # search_resultsKujiR23 = ["Kuji R 2023", values_list_KujiAD2023]
                                search_results_item.append(values_list_KujiR2023)

                            # Define query to search Kuji R 2023 table for null price matches:
                            query_KujiR23_Null = "SELECT DPRFQ, DATE FROM `Kuji R Storage 2023` WHERE (`Parts Number`=\'" + str(
                                listCombo[i][0]) + "\'"
                            # Amend SQL query to include search for NSN numbers if applicable for given PN:
                            if listCombo[i][1] != "No NSN entered" and listCombo[i][1] \
                                    != "Invalid NSN of format PN[PartNumber] entered":
                                query_KujiR23_Null = query_KujiR23_Null + " OR `NSN`=\'" + str(
                                    listCombo[i][1]).rstrip() + "\'"
                            # Perform query to search 4AD Max 2023 Storage table for matches:
                            query_KujiR23_Null = query_KujiR23_Null + ") AND ALMCQ IS NOT NULL"
                            # print(query_KujiR23_Null)
                            """
                            # Old obselete code for searching for and returning matches:
                            df = pd.read_sql(query_KujiR23_Null, conn)
                            # Format DATE column results to only show RFQ issue date (May need to revise this later-MW)
                            # print(df['DATE'].str.find("-"))
                            # print(df['DATE'].str[:7])
                            df['DATE'] = df['DATE'].str[:7]
                            # Store query results as a list of values
                            values_list_KujiR2023_Null = df.values.tolist()
                            # print(values_list_KujiR2023_Null)
                            # Confirm that matching values were found in the table:
                            if len(values_list_KujiR2023_Null) != 0:
                                # search_resultsKujiR23 = ["Kuji R 2023", values_list_KujiAD2023]
                                search_results_item.append(values_list_KujiR2023_Null)
                            """
                            # Experimental code to condense and reformat row strings:
                            result = cursor.execute(query_KujiR23_Null)
                            cursor_KujiR23_Null_RList = []
                            for row in result:
                                # print(row.DPRFQ)
                                dprfq_Mod = row.DPRFQ + " (2023)"
                                # print(dprfq_Mod)
                                if cursor_KujiR23_Null_RList.count(dprfq_Mod + ", No bid") == 0:
                                    cursor_KujiR23_Null_RList.append(dprfq_Mod + ", No bid")
                            if len(cursor_KujiR23_Null_RList) != 0:
                                # print(cursor_KujiR23_Null_RList)
                                search_results_item.append(cursor_KujiR23_Null_RList)

                            # Define query to search Kuji Numbered 2023 table:
                            query_KujiNum23 = "SELECT DPRFQ, DATE, ALMCQ, SUGGEST, NOTES, `Company Quotes` FROM `Kuji Numbered 2023` WHERE (`Parts Number`=\'" + str(
                                listCombo[i][0]) + "\'"
                            # Amend SQL query to include search for NSN numbers if applicable for given PN:
                            if listCombo[i][1] != "No NSN entered" and listCombo[i][1] \
                                    != "Invalid NSN of format PN[PartNumber] entered":
                                query_KujiNum23 = query_KujiNum23 + " OR `NSN`=\'" + str(
                                    listCombo[i][1]).rstrip() + "\'"
                            # Perform query to search 4AD Max 2023 Storage table for matches:
                            query_KujiNum23 = query_KujiNum23 + ") AND ALMCQ IS NOT NULL"
                            # print(query_KujiNum23)
                            df = pd.read_sql(query_KujiNum23, conn)
                            # Format DATE column results to only show RFQ issue date (May need to revise this later-MW)
                            # print(df['DATE'].str.find("-"))
                            # print(df['DATE'].str[:7])
                            df['DATE'] = df['DATE'].str[:7]
                            # Format ALMCQ and SUGGEST column results to read as currency values of form '$x.00'
                            df['ALMCQ'] = df['ALMCQ'].apply(lambda x: "${:.2f}".format((x)))
                            df['SUGGEST'] = df['SUGGEST'].apply(lambda x: "${:.2f}".format((x)))
                            # Store query results as a list of values
                            values_list_KujiNum2023 = df.values.tolist()
                            # print(values_list_KujiNum2023)
                            # Confirm that matching values were found in the table:
                            if len(values_list_KujiNum2023) != 0:
                                # search_resultsKujiNum23 = ["Kuji Numbered 2023", values_list_KujiNum2023]
                                search_results_item.append(values_list_KujiNum2023)

                            # Define query to search Kuji Numbered 2023 table for Null price Matches:
                            query_KujiNum23_Null = "SELECT DPRFQ, DATE FROM `Kuji Numbered 2023` WHERE (`Parts Number`=\'" + str(
                                listCombo[i][0]) + "\'"
                            # Amend SQL query to include search for NSN numbers if applicable for given PN:
                            if listCombo[i][1] != "No NSN entered" and listCombo[i][1] \
                                    != "Invalid NSN of format PN[PartNumber] entered":
                                query_KujiNum23_Null = query_KujiNum23_Null + " OR `NSN`=\'" + str(
                                    listCombo[i][1]).rstrip() + "\'"
                            # Perform query to search 4AD Max 2023 Storage table for matches:
                            query_KujiNum23_Null = query_KujiNum23_Null + ") AND ALMCQ IS NULL"
                            # print(query_KujiNum23_Null)
                            """
                            # Obsolete code for searching for and returning null matches
                            df = pd.read_sql(query_KujiNum23_Null, conn)
                            # Format DATE column results to only show RFQ issue date (May need to revise this later-MW)
                            # print(df['DATE'].str.find("-"))
                            # print(df['DATE'].str[:7])
                            df['DATE'] = df['DATE'].str[:7]
                            # Store query results as a list of values
                            values_list_KujiNum2023_Null = df.values.tolist()
                            # print(values_list_KujiNum2023_Null)
                            # Confirm that matching values were found in the table:
                            if len(values_list_KujiNum2023_Null) != 0:
                                # search_resultsKujiNum23_Null = ["Kuji Numbered 2023", values_list_KujiNum2023_Null]
                                search_results_item.append(values_list_KujiNum2023_Null)
                            """
                            # Experimental code to condense and reformat row strings:
                            result = cursor.execute(query_KujiNum23_Null)
                            cursor_KujiNum23_Null_RList = []
                            for row in result:
                                # print(row.DPRFQ)
                                dprfq_Mod = row.DPRFQ + " (2023)"
                                # print(dprfq_Mod)
                                if cursor_KujiNum23_Null_RList.count(dprfq_Mod + ", No bid") == 0:
                                    cursor_KujiNum23_Null_RList.append(dprfq_Mod + ", No bid")
                            if len(cursor_KujiNum23_Null_RList) != 0:
                                # print(cursor_KujiNum23_Null_RList)
                                search_results_item.append(cursor_KujiNum23_Null_RList)

                            # Confirm that master storage list for item is not empty:
                            if (len(search_results_item)) != 0:
                                search_results_item_string = " / ".join(str(e) for e in search_results_item)
                                # print(search_results_item_string)
                                search_results_item_string = search_results_item_string.replace("[", "")
                                search_results_item_string = search_results_item_string.replace("]", "")
                                search_results_item_string = search_results_item_string.replace('\'', "")
                                # print(search_results_item_string)
                                # Add resulting findings of table searches to
                                search_results.append((entry_value, search_results_item_string))
                                print(str(search_results))
                            else:
                                # If no matches for any qualifying values where found, return a "No matches found in Database" message to user.
                                search_results.append((entry_value, "No matches found in Database"))
                                # print(str(search_results))

                        # Return and display results of search to the run screen:
                        iterator = 1
                        # search_Values = []
                        for i in search_results:
                            print(str(iterator) + "> " + str(i))
                            # print(i[1])
                            # search_Values.append(i[1])
                            iterator += 1
                            # print(search_Values)

                        if 0 < len(search_results) <= 100:
                            # Create new window to display results:
                            new_Window = Toplevel(root)
                            new_Window.geometry("800x200")
                            new_Window.title("Search Results")
                            # Create a Canvas:
                            my_canvas = Canvas(new_Window)
                            my_canvas.pack(side=LEFT, fill=BOTH, expand=1)
                            # Add a Scrollbar to the Canvas:
                            my_scrollbar = ttk.Scrollbar(new_Window, orient=VERTICAL, command=my_canvas.yview)
                            my_scrollbar.pack(side=RIGHT, fill=Y)
                            # Configure the Canvas:
                            my_canvas.configure(yscrollcommand=my_scrollbar.set)
                            my_canvas.bind('<Configure>',
                                   lambda e: my_canvas.configure(scrollregion=my_canvas.bbox("all")))
                            # Create a Frame inside the Canvas
                            secondFrame = Frame(my_canvas)
                            # Add that new Frame to a Window in the Canvas:
                            my_canvas.create_window((0, 0), window=secondFrame, anchor="nw")
                            # Define first column of new window for Part Numbers:
                            labelColumn1 = Label(secondFrame, text="Part Number:", width=23, bg="white",
                                         relief="solid")
                            labelColumn1.grid(row=0, column=0, padx=2, pady=2)
                            # Define second column of new window for NSN Numbers:
                            labelColumn2 = Label(secondFrame, text="NSN Number:", width=33, bg="white",
                                             relief="solid")
                            labelColumn2.grid(row=0, column=1, padx=2, pady=2)
                            # Define Third column of new window for search results for Master past QUOTE DATABASE table:
                            labelColumn3 = Label(secondFrame, text="Results:", width=43, bg="white", relief="solid")
                            labelColumn3.grid(row=0, column=2, padx=2, pady=2)
                            """
                            # Define Fourth column of new window for search results for 2AD 9-2 HZ table:
                            labelColumn4 = Label(secondFrame, text="2AD 9-2 HZ:", width=43, bg="white", relief="solid")
                            labelColumn4.grid(row=0, column=4, padx=2, pady=2)
                            """
                            for x in range(len(search_results)):
                                # Obtain tuple containing current part number and NSN Number in list:
                                temp_str = search_results[x][0]
                                # print(temp_str)
                                height_value = 2  # Defining base height for text boxes
                                # Confirm that matches for current PN and/or NSN were found in DB:
                                if not str(search_results[x][1]).startswith("No"):
                                    height_value = 4  # If a result was found, adjust row text box height for uniformity
                                if temp_str.find("(") != -1:
                                    # Storage variables for start and end indexes of part number in tuple turned string:
                                    part_num_start_index = temp_str.index("(") + 2
                                    part_num_end_index = temp_str.index(",") - 1
                                    # print(temp_str[part_num_start_index:part_num_end_index])
                                    # Store current part number in string variable:
                                    part_num = temp_str[part_num_start_index:part_num_end_index]
                                    # Storage variables for start and end indexes of NSN number in tuple turned string:
                                    nsn_num_start_index = temp_str.index(",") + 3
                                    nsn_num_end_index = temp_str.index(")") - 1
                                    # print(temp_str[nsn_num_start_index:nsn_num_end_index])
                                    # Store current NSN number in string variable:
                                    nsn_number = temp_str[nsn_num_start_index:nsn_num_end_index]
                                    # Create new Text area widget to store part number in:
                                    part_area = Text(secondFrame, width=20, height=height_value, relief="solid")
                                    # Insert part number into new text area:
                                    part_area.insert(INSERT, part_num)
                                    # Insert Text area into appropriate row and column in the new window's grid:
                                    part_area.grid(row=x + 1, column=0, padx=2, pady=2)
                                    # Create new Text area widget to store NSN number in:
                                    nsn_area = Text(secondFrame, width=30, height=height_value, relief="solid")
                                    # Insert NSN number into new text area:
                                    nsn_area.insert(INSERT, nsn_number)
                                    # Insert Text area into appropriate row and column in the new window's grid:
                                    nsn_area.grid(row=x + 1, column=1, padx=2, pady=2)
                                    # Create new Text area widget to store search results in:
                                    results_area = Text(secondFrame, width=40, height=height_value, relief="solid")
                                    if not str(search_results[x][1]).startswith("No"):
                                        # Remove square brackets from search results:
                                        search_results_string = search_results[x][1].replace("[", "")
                                        search_results_string = search_results_string.replace("]", "")
                                        # print(search_results_string)
                                        # Insert search results into new text area:
                                        results_area.insert(INSERT, search_results_string)
                                    else:
                                        results_area.insert(INSERT, str(search_results[x][1]))
                                    # Insert Text area into appropriate row and column in the new window's grid:
                                    results_area.grid(row=x + 1, column=2, padx=2, pady=2)
                                    # If a match for the part number or NSN had been found in Database, add a scrollbar:
                                    if not str(search_results[x][1]).startswith("No"):
                                        # Define a linked scrollbar to help users scroll through search results text area.
                                        resultsHsb = Scrollbar(secondFrame, orient="vertical",
                                                       command=results_area.yview)
                                        # Insert scrollbar next to search results text area
                                        resultsHsb.grid(row=x + 1, column=3)
                                        # Configure scrollbar
                                        results_area.configure(yscrollcommand=resultsHsb.set)
                                    """
                                    else:
                                        # Create new Text area widget to store part number in:
                                        part_area = Text(secondFrame, width=20, height=height_value, relief="solid")
                                        # Insert part number into new text area:
                                        part_area.insert(INSERT, temp_str)
                                        # Insert Text area into appropriate row and column in the new window's grid:
                                        part_area.grid(row=x + 1, column=0, padx=2, pady=2)
                                        # Create new Text area widget to store NSN number in:
                                        nsn_area = Text(secondFrame, width=30, height=height_value, relief="solid")
                                        # Insert NSN number into new text area:
                                        nsn_area.insert(INSERT, "No valid NSN Number entered in excel file.")
                                        # Insert Text area into appropriate row and column in the new window's grid:
                                        nsn_area.grid(row=x + 1, column=1, padx=2, pady=2)
                                        # Create new Text area widget to store search results in:
                                        results_area = Text(secondFrame, width=40, height=height_value, relief="solid")
                                        # Insert search results into new text area:
                                        results_area.insert(INSERT, str(search_results[x][1]))
                                        # Insert Text area into appropriate row and column in the new window's grid:
                                        results_area.grid(row=x + 1, column=2, padx=2, pady=2)
                                    """
                                    # If a match for the part number or NSN had been found in Database, add a scrollbar:
                                    if not str(search_results[x][1]).startswith("No"):
                                        # Define a linked scrollbar to help users scroll through search results text area.
                                        resultsHsb = Scrollbar(secondFrame, orient="vertical",
                                                       command=results_area.yview)
                                        # Insert scrollbar next to search results text area
                                        resultsHsb.grid(row=x + 1, column=3)
                                        # Configure scrollbar
                                        results_area.configure(yscrollcommand=resultsHsb.set)
                                else:
                                    # print("No matches")
                                    # Create new Text area widget to store part number in:
                                    part_area = Text(secondFrame, width=20, height=height_value, relief="solid")
                                    # Insert part number into new text area:
                                    part_area.insert(INSERT, temp_str)
                                    # Insert Text area into appropriate row and column in the new window's grid:
                                    part_area.grid(row=x + 1, column=0, padx=2, pady=2)
                                    # Create new Text area widget to store NSN number in:
                                    nsn_area = Text(secondFrame, width=30, height=height_value, relief="solid")
                                    # Insert NSN number into new text area:
                                    nsn_area.insert(INSERT, "No valid NSN Number entered in excel file.")
                                    # Insert Text area into appropriate row and column in the new window's grid:
                                    nsn_area.grid(row=x + 1, column=1, padx=2, pady=2)
                                    # Create new Text area widget to store search results in:
                                    results_area = Text(secondFrame, width=40, height=height_value, relief="solid")
                                    # Insert search results into new text area:
                                    results_area.insert(INSERT, str(search_results[x][1]))
                                    # Insert Text area into appropriate row and column in the new window's grid:
                                    results_area.grid(row=x + 1, column=2, padx=2, pady=2)
                                    # If a match for the part number or NSN had been found in Database, add a scrollbar:
                                    if not str(search_results[x][1]).startswith("No"):
                                        # Define a linked scrollbar to help users scroll through search results text area.
                                        resultsHsb = Scrollbar(secondFrame, orient="vertical", command=results_area.yview)
                                        # Insert scrollbar next to search results text area
                                        resultsHsb.grid(row=x + 1, column=3)
                                        # Configure scrollbar
                                        results_area.configure(yscrollcommand=resultsHsb.set)
                        elif len(search_results) > 100:
                            # Code designed to export the results of particularly large files to a new sheet for the Excel file.
                            df1 = pd.DataFrame(search_results, columns=["Parts and NSN Numbers", "Search Results"])
                            # print(df1)
                            try:
                                with pd.ExcelWriter(searchLocation, mode='a', if_sheet_exists="replace") as writer:
                                    df1.to_excel(writer, sheet_name='Raw DB Search Results')
                            except FileNotFoundError as error_A:
                                print(error_A)
                            # Create Window to explain that results were exported to Excel due to size of file:
                            new_Window_O = Toplevel(root)
                            # Define size of new Window:
                            new_Window_O.geometry("400x200")
                            # Set title of new window:
                            new_Window_O.title("Search Results")
                            # Create a frame in the new window to store and organize window's contents:
                            overflow_frame = Frame(new_Window_O)
                            # Insert frame into window:
                            overflow_frame.pack()
                            # Create a text label to inform user of where to find search results:
                            overflow_message = tkinter.Label(overflow_frame,
                                                     text="Due to size of file, results were exported to a new sheet named \'Raw DB Search Results\' created in given file for permanent user review",
                                                     width=100)
                            # Pack text label into frame:
                            overflow_message.pack()
                except pyodbc.Error as error:
                    # If an error occurred during the connection process, report the error to the user:
                    print(str(error))
                    error_string.set(str(error))
            except ValueError as error:
                print(str(error))
                error_string.set(str(error))
        except FileNotFoundError as error:
            # If file was not found, report the error to the user:
            print("File not found." + str(error))
            error_message = "File not found." + str(error)
            error_string.set(error_message)
    else:
        error_string.set("Please enter the location of an Excel file.")


"""
# Test frame swap functions for when user inputs Access database path.
def swap_access():
    # Swap to frame for requesting Access database information
    access_frame.pack(fill='both', expand=1)
    excel_frame.pack_forget()


def swap_excel():
    # Swap back to initial frame.
    excel_frame.pack(fill='both', expand=1)
    access_frame.pack_forget()
"""

search_txt = Label(excel_frame, text="Enter location of Excel file in system:")
# Get the location of the file in the current computer directory from the user:
input_txt = Entry(excel_frame, width=40)
submit_button = Button(excel_frame, text="Submit", command=databaseSearch)
error_txt = Label(excel_frame, textvariable=error_string)
excel_frame.pack()
search_txt.pack()
input_txt.pack()
submit_button.pack()
error_txt.pack()
"""
# Defining contents of frame for requesting Access database information:
access_path_txt = Label(access_frame, text="Enter location of Access file in system:")
# Get the location of the file in the current computer directory from the user:
access_path_input = Entry(access_frame, width=40)
password_txt = Label(access_frame, text="Enter password to system if applicable:")
# Get the password for the database from the user:
password_input = Entry(access_frame, width=40)
access_submit = Button(access_frame, text="Submit", command=databaseSearch)
error_txt_b = Label(access_frame, text_variable=error_string)
access_path_txt.pack()
access_path_input.pack()
access_submit.pack()
error_txt_b.pack()
"""
mainloop()
