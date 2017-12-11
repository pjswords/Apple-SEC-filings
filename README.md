# Apple-SEC-filings
These Python scripts retrieve and parse Apple's annual SEC filings from 2007 
through 2015. Apple-SEC-filings-py2.py works with Python 2, and 
Apple-SEC-filings-py3.py works with Python 3. Both scripts perform 
the same function.

After retrieving a report from Apple's website, the script 
searches for the section titled "Item 1A. Risk Factors" and builds a 
dictionary from the words in that section. The script then places 
the contents of the dictionary into an object that corresponds to 
a Microsoft Excel worksheet. While building the worksheets, the script also
provides console output showing which words have the highest frequency.
Once a worksheet is built for each report, the script saves them in a single
XLSX file that can be used with Tableau or other analytics software.

Note: These scripts require access to the BeautifulSoup, OpenPyXL, 
and nltk libraries.
