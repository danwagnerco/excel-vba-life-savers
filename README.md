Excel VBA Life-Savers
=====================

### After running a script, Excel will not open

You might be lucky enough to get a "Missing Project or Library" error from Excel, but it's very possible that when you click on an Excel file it simply goes unresponsive and quits. Bummer!

[User Alulla on Ozgrid](http://www.ozgrid.com/forum/showthread.php?t=179860) has a solution that works like gangbusters. Here it is:

1. Go to the Developer Tab and click “Macro Security”.
2. Click the dot entitled “Disable all macros with notification”
3. Go to Trusted Locations (on the left) and temporarily click “Disable all Trusted Locations”
4. Go to Trusted Documents and temporarily click “Disable all Trusted Documents”
5. Click Ok and now open “Allocations” or whatever file was not working
6. Do not click “Enable Macros”…instead go to the Developer Tab and open Visual Basic
7. On Visual Basic, click save and then click Debug > Compile VBAProject
8. Save in VB and Save in the Excel spreadsheet & then close the file
9. Open it and Enable Macros. Should work now.

Issue occurrences:
- 12/16/2014 (initial)
- 1/27/2015 (after Stephen pasted in data while I was working remote in SF for Symposium)
