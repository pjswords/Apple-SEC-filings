# Apple-SEC-filings
This program retrieves and parses Apple's annual SEC filings from 2007 
through 2015. After retrieving a report from Apple's website, the program 
searches for the section titled "Item 1A. Risk Factors" and builds a 
dictionary from the words in that section. The program then places 
the contents of the dictionary into an object that corresponds to 
a Microsoft Excel worksheet. While building the worksheets, the program also
provides console output showing which words have the highest frequency.
Once a worksheet is built for each report, the program saves them in a single
XLSX file that can be used with Tableau or other analytics software.
