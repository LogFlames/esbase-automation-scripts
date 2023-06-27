# esbase-automation-scripts

A script for changing multiple accessions in the PHP framework of ESBase.

# Usage

## First time setup

Install [python](https://www.python.org/) and [git](https://gitforwindows.org/).

Open Git Bash, navigate to the folder you wish to place the scripts in using `cd` to change directory, `ls` to see all files and folders in the current directory and `pwd` to print the current working directory (where you are). Then run

```bash
git clone https://github.com/logflames/esbase-automation-scripts.git
cd ./esbase-automation-scripts/
python -m pip install -r requirements.txt
```

## Updating

Navigate to the `esbase-automation-scripts/` folder. Inside the folder run the following commands:
```bash
git pull
python -m pip install -r requirements.txt
```

## Running

Make a copy of the `example.xlsx` excel-file. This can be done with (change `name_of_new_file` to something of your chosing):
```bash
cp example.xlsx name_of_new_file.xlsx
```

Go to [http://esbase.nrm.se](http://esbase.nrm.se) and get all the `id`:s of the fields you wish to change using the `Inspect` option in the browser.
For fields with search options, make sure to use the drop-down select as the script doesn't wait for the search to finish (person, locality, etc.).

For select fields you can enter either the visible text or the `value` on the option. The script will try to figure out which one you want based on what options are available. If your new value only exists in one of the groups (either the values or the visible texts) it will select based on that group. If you new value exists in both groups it will select based on the value. If multiple options have the same visible text, make sure to select based on value.

Once the excel-sheet is filled in. Run the following command (make sure to replace `name_of_new_file` with the name you chose earlier):
```bash
python set_property_for_accnrs.py --excel name_of_new_file.xlsx
```

# Contact

If there are any bugs or feature-requests, [mail me](mailto:elilun03@gmail.com) or [create an issue](https://github.com/LogFlames/esbase-automation-scripts/issues/new).

Made by: Elias Lundell

