# TD-Webbroker-investment-calculator
A python script which reads investment figures from excel documents and calculates where to invest given a dollar amount.

The calculator follows the Canadian Couch Potato model and essentially calculates the money you need to add to your holdings to achieve the desired proportions of the 4 holdings in the CCP portfolio.

Usage:
To use, save all files from the repository to the same folder, as well as the exported account information from TD Webbroker as CAD_CASH and CAD_TFSA excel files. From the command line, use './calc_investments.sh (AMOUNT_INVESTING) (INVEST_IN_TFSA)'.

If the CAD_CASH or CAD_TFSA files are not found, it will simply use the data from the most recent entry in the investments file.

Required libraries:
Openpyxl ('pip3 install openpyxl' to install)
