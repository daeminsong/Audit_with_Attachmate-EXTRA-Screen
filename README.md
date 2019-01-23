
Obviously this macro is for a very specific purpose. 

What it does – 

1.	Creates worksheets based on a list of strings in the worksheet “List”
2.	Creates an index sheet so that auditors could easily navigate through the workbook
3.	Loops through the strings – takes a screenshot(s) from Attachmate - EXTRA!, and saves it (or them) under the corresponding spreadsheet by order. 

Requirements – 

Macro-enabled workbook (.xlsm)
Worksheet named “List” – of course you can modify it if you would like
Attachmate - EXTRA! (and you need to be signed in so that VBA could interact with it)

Note – 

* You would likely want to modify line 69, so that VBA could see the pages that you want to take screenshots. 
* Nothing should be over the Attachmate - EXTRA! window, it won’t make proper screenshots. VBA won’t recognize the Attachmate - EXTRA! as a separate program per se - It basically recognize the size and location of the windows on the screen. It probably is easier to think it as cropping a part of screen. 
