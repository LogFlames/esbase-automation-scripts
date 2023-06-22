from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
import openpyxl
import argparse
import re
import time

def main():
    parser = argparse.ArgumentParser(
            prog = "Esbase Automation Script",
            description = "Automates interactions with the php script using selenium")
    parser.add_argument("--excel", dest = "excel", help = "The path to an excel file containing three columns, with the names: 'accnr', 'property_id', 'new_value'", required = True)

    args = parser.parse_args()

    wb = openpyxl.load_workbook(args.excel).active
    if wb.max_column % 2 != 1:
        raise Exception("The excel file must have an odd number of columns")

    to_change = []
    for row in range(0, wb.max_row):
        values = []
        for col in wb.iter_cols(1, wb.max_column):
            val = col[row].value
            if val is None:
                val = ""
            val = str(val)
            values.append(val)
        if row == 0:
            if values[0] != "accnr":
                raise Exception("The first column in the Excel-file must be accnr")
            for prop in range(1, len(values), 2):
                if "property_id" not in values[prop]:
                    raise Exception(f"The column {prop + 1} must contain 'property_id'")
            for nv in range(2, len(values), 2):
                if "new_value" not in values[nv]:
                    raise Exception(f"The column {nv + 1} must contain 'new_value'")
        else:
            if re.match("[1-9]{1}[0-9]{8}", values[0]) is None:
                raise Exception(f"Invalid accnr '{values[0]}', please enter on the form in the database: eg 'A2022/12345' -> '202212345', 'B2022/12345' -> '102212345'")

            to_change.append({ "accnr": values[0], "property_id": values[1::2], "new_value": values[2::2] })
            print(to_change[-1])

    opt = webdriver.FirefoxOptions()
    opt.binary_location = "./FirefoxPortable/App/Firefox64/firefox.exe"

    driver = webdriver.Firefox(options = opt)
    driver.implicitly_wait(15)
    driver.get("http://esbase.nrm.se")

    input(" -- Press enter when you've logged in: -- \n")


    first = True
    start = time.time()

    with open("log.txt", "a") as f:
        f.write(f"Run started at {start}" + "\n")
    with open("log.csv", "a") as f:
        f.write(f"accnr,property_id,old_value,new_value" + "\n")

    for i, row in enumerate(to_change):
        driver.get("http://esbase.nrm.se/accession?id=" + row["accnr"])

        elem_first = None
        old_values = []
        for j in range(len(row["property_id"])):
            try:
                elem = driver.find_element(By.ID, row["property_id"][j])
                if j == 0:
                    elem_first = elem
            except NoSuchElementException:
                raise Exception(f"Invalid property_id '{row['property_id'][j]}', could not find element with id.")

            if elem.tag_name == "select":
                s = Select(elem)
                old_values.append(s.first_selected_option.text)
                s.select_by_visible_text(row["new_value"][j])
            else:
                old_values.append(elem.text)
                elem.clear()
                elem.send_keys(row["new_value"][j])

        if first:
            ans = input("Does everything look good? (y/n) ")
            while ans not in ["y", "n"]:
                ans = input("Does everything look good? (y/n) ")
            if ans != "y":
                print("Aborting operations...")
                with open("log.txt", "a") as f:
                    f.write("Run aborted" + "\n")
                driver.close()
                exit()
            first = False
            start = time.time()

        if elem_first is None:
            raise Exception("First element was None, cannot submit")
        else:
            elem_first.submit()

        log_txt = f"Updated '{row['accnr']}': '{row['property_id']}' from '{old_values}' to '{row['new_value']}'"
        log_csv = f"{row['accnr']},{row['property_id']},'{old_values}','{row['new_value']}'"
        now = time.time()
        eta = int((now - start) * len(to_change) / (i + 1))
        print(log_txt + "\tPassed time/Est total\t" + f"{int(now - start) // 60}:{int(now - start) % 60:02}/{eta//60}:{eta % 60:02}")
        with open("log.txt", "a") as f:
            f.write(log_txt + "\n")
        with open("log.csv", "a") as f:
            f.write(log_csv + "\n")

    driver.close()

if __name__ == "__main__":
    main()
