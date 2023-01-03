# NessusReportGenerator

Python script which can:
1. Filter Unwanted False positives and remove duplicates of Vulnerabilities.
2. Background color for Vulnerabilities with Purple as Critical, Red as High and Yellow for Low
3.Output name is dependent on input name, file type is changed from csv to xlsx
4.Converts all csv files in the current directory only.


NOTE:
1.USE AT YOUR OWN RISK, may be susceptible to bugs, no issues as far as i have tested.
2.Report generation on Nessus need to have only the following Columns: Host, Name, Description, Risk, Synopsis, Solution, CVE and CVSSv3.0 .
3.Any other Header value in the Nessus report is not currently compatible( Such as Plugin ID). Please report any bugs.
4. A LIST OF SOME COMMON VULNERABILITIES HAVE BEEN ADDED  FOR FILTERING. PLEASE REMOVE THIS.
