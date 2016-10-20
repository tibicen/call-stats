#Call stats 
##[pylab][openpyxl][datetime]
####Description:
Little script for statistical data from monthly detailed reports of my phone delivery.

For now it only works with reports from:
- T-mobile Poland

####Usage:
Download all your nessesary reports in txt format. Put it in folder named by your telephone number. Then run the script.

####Data:
You will get excel file with average, min, max usage of MB, call minutes and SMS during the period of all your reports.
In other sheets you will get all your calls joined and sorted by date.
You will also get jpg files with MB usage during the month, fit curve and estimation.

####TODO:
- expand it for more telephone companies (i need example reports, for scraping data);
- automate the process for sorting reports by user/number
- create exe.

<img src="https://github.com/tibicen/call-stats/blob/master/examples/yearly-MIN-SMS-MB.png" width="300">
<img src="https://github.com/tibicen/call-stats/blob/master/examples/monthly-MB.png" width="300">
<img src="https://github.com/tibicen/call-stats/blob/master/examples/raport2016.7.21.jpg" width="300">
