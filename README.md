# dav-report-rkumarm2
DAV report automation

Six spreadsheet files with the name: 
SEBH_DetailedDeviceDetails_DDMmmYYYY ,
SEBH_ViewManagedDevices_DDMmmYYYY ,
SECH_DetailedDeviceDetails_DDMmmYYYY ,
SECH_ViewManagedDevices_DDMmmYYYY ,
SETW_DetailedDeviceDetails_DDMmmYYYY ,
SETW_ViewManagedDevices_DDMmmYYYY 

needs to be given as input.(The date, month in which these six reports were generated will be referred as generated date, generated month in the below description).

In the first text box, one needs to enter the date as given in generated date.

The generated month will always be considered, one needs to type the number of previous months to be considered from the generated month. 

If August is generated month, and May, June, July needs to be considered along with August, then 3 should be entered in months input text box.

The output file will be downloaded in the "C:\Users\"username" folder.


After executing the script for once, can remove the "pip install ..." lines in the code to avoid unnecessary time in code execution.
