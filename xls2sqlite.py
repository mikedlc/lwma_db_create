# Extra packages may be required, if so run these:
#$ pip install --upgrade xlrd
#$ pip install --upgrade cutplace

# Table names cannot begin with numbers(?) so make sure they begin with a letter. Using 'USB_<date>.csv' for now.
# Need to add logic to check for existing tables and not just assume DROP if one exists already...

# Instructions:
# Put this xls2sqlite.py in a directory with .xls output files from GDrive, and run it
# 'python .\csv2sqlite.py dbOut.db' for example.
# Had issues figuring out encoding for Kamstrup .csv output, so used their XLS output converted to .csv by python for standard encoding which worked.
# Program will create separate formatted .csv's for each .xls and will create a table in the db for each.

from __future__ import print_function
import sqlite3
import csv
import os
import glob
import sys
import pdb
import xlrd
import shutil


print("\n\n\n\n*** Converting XLS files found in working dir to CSV***\n")

for xlsfile in glob.glob(os.path.join(os.getcwd(), "*.xls")):
    # Print which file is being worked on...
    print("\nConverting: " + xlsfile)
    
    with xlrd.open_workbook(xlsfile, logfile=open(os.devnull, 'w')) as wb: #null log file to suppress unimportant warnings
        sh = wb.sheet_by_index(0) 
        csvfile = os.path.splitext(os.path.basename(xlsfile))[0] + ".csv"
        # Write .csv content
        with open(csvfile, 'w', newline="") as f:
            c = csv.writer(f)
            for r in range(sh.nrows):
                c.writerow(sh.row_values(r))


    #Kamstrup headers are sloppy and sqlite gets confused. Fixing them manually. This can probably be done cleaner but this works...
    from_file = open(csvfile)
    to_file = open(os.path.splitext(os.path.basename(csvfile))[0] + "_formatted.csv", 'w', newline="")
    to_file.write("Serial_number,Name,Meter_type,Consumption_type,Volume_V1,Receive_time,Volume_H,Operating_hour_counter,Minimum_flow_temperature_H,Minimum_external_temperature_H,Info,Avr_ext_temp_H\n")
    writer = csv.writer(to_file)
    # write the meter readings in under the clean headers.
    for row in csv.reader(from_file):
        if row[0]!="Serial number":
            writer.writerow(row)
    from_file.close()
    # Remove .csv with dirty headers
    os.remove(csvfile)
    to_file.close()
    print("Converted.")

# Done creating .csv's... on to importing them to DB

print("\n\n*** Importing generated CSVs into sqlite ***\n")
db = sys.argv[1] #first argument is the name of the output db

conn = sqlite3.connect(db)
conn.text_factory = str  # allows utf-8 data to be stored

c = conn.cursor()

# Process each .csv found in the working directory
for csvfile in glob.glob(os.path.join(os.getcwd(), "*.csv")):
    
    # remove the path and extension and use what's left as a table name
    tablename = os.path.splitext(os.path.basename(csvfile))[0]
	
	# Print .csv being processed and the resulting table name
    print("\nImporting: " + csvfile)
    print("Table Name: " + tablename)

    # open the csv for processing
    with open(csvfile, "r") as f:
        reader = csv.reader(f, delimiter=',')

        header = True
        
        #for each row...
        for row in reader:
            
            if header:
                
                header = False
                
                #delete the table if it exists... this logic can be tweaked.
                sql = "DROP TABLE IF EXISTS %s" % tablename
                c.execute(sql)
                # Create the table for this reading session
                sql = "CREATE TABLE %s (%s)" % (tablename, ", ".join([ "%s" % column for column in row ]))
                #sql = CREATE TABLE DLC20180712 (Serial_number, Name, Meter_type, Consumption_type, Volume_V1, Receive_time, Volume_H, Operating_hour_counter, Minimum_flow_temperature_H, Minimum_external_temperature_H, Info, Avr_ext_temp_H)
                #print("\n*** sql: " + sql + "\n")
                c.execute(sql)

                for column in row:
                    if column.lower().endswith("_id"):
                        index = "%s__%s" % ( tablename, column )
                        sql = "CREATE INDEX %s on %s (%s)" % ( index, tablename, column )
                        c.execute(sql)

                insertsql = "INSERT INTO %s VALUES (%s)" % (tablename,
                            ", ".join([ "?" for column in row ]))

                rowlen = len(row)
            else:
                # skip lines that don't have the right number of columns
                #print("SKIP")
                if len(row) == rowlen:
                    c.execute(insertsql, row)

        conn.commit()

print("\n\n\nComplete: " + db + " Contains a table for each reading session .xls found in the working dir.\n\n\n")

c.close()
conn.close()
