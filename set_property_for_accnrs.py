from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import csv
import argparse
import re

def main():
    parser = argparse.ArgumentParser(
            prog = "Esbase Automation Script",
            description = "Automates interactions with the php script using selenium")
    parser.add_argument("--csv", dest = "csv", help = "The path to a csv file containing three columns, with the names: 'accnr', 'property_id', 'new_value'", required = True)

    args = parser.parse_args()

    with open(args.csv) as csvfile:
        reader = csv.reader(csvfile)
        rows = [*reader]
        if len(rows) == 0:
            raise Exception("The csv file must contain atleast one row")
        header = rows[0]
        if header != ["accnr", "property_id", "new_value"]:
            raise Exception("The CSV-file must contain the headers 'accnr', 'property_id', 'new_value'")

        to_change = []        
        for row in rows[1:]:
            if not re.match("[1-9]{1}[0-9]{8}", row[0]):
                raise Exception("Invalid accnr, please enter on the form in the database: eg 'A2022/12345' -> '202212345', 'B2022/12345' -> '102212345'")
            to_change.append({ "accnr": row[0], "property_id": row[1], "new_value": row[2] })

    driver = webdriver.Firefox()
    driver.implicitly_wait(15)
    driver.get("http://esbase.nrm.se")

    input("Press enter when you've logged in: ")

    for row in to_change:
        driver.get("http://esbase.nrm.se/accession?id=" + row["accnr"])
        try:
            elem = driver.find_element(By.ID, row["property_id"])
        except NoSuchElementException:
            raise Exception(f"Invalid property_id '{row['property_id']}', could not find element with id.")

        elem.clear()
        elem.send_keys(row["new_value"])

        elem.submit()

    driver.close()

if __name__ == "__main__":
    main()



