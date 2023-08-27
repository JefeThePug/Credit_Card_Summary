# Credit_Card_Summary
 Credit Card Summary Web Scraping tool

Scraping from my bank's website, Credit Card Purchases are collected for a given month.  

These purchases are then organised into categories based on a dictionary with keys 
that have been manually inserted based on what is written in the credit card statement
to sort the purchases into types of spending.  New/unique purchases not recognised by
the dictionary will remain without a category and can be changed or updated in the 
dictionary later as required.

The purchases with categories are then built into an Excel file with VBA to allow me
to select the month I would like to view.  Reading from the sheet created by the .py 
file, it displays the summary of purchases on a new sheet.  Purchases without a 
category can be updated manually here and the summary will adjust itself.
Also included in the summary are unchanging fees (like rent and internet) as well as 
changing fees (which will have to be manually added) to provide a summary of spending 
for the entire month.

Further VBA allows for a button to click when the summary has finished being adjusted.
It will then automatically save the summary page as a PDF and email that PDF to the
recipient chosen within the VBA code. 
