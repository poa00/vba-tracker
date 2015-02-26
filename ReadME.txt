Excel VBA Macro- Project Tracking and reporting
===============================================

This macro application was created for large Company with many small projects from clients, 
to efficiently track each projects status , each employees status and the resource level status like whether access is still granted or revoked

requirement of company:
=> Company had many project and frequent on-boarding/off-boarding in resource pool
=> each project had to request Application/Database/tools access at start and had to be revoked once project is over/resource off-boarding
=> if this is not done , SLA will be broken and had to pay the penalties

As a Solution,
I consolidate all the data in common format in single workbook and written few macros to generate following reports
1. User level
2. Over all Report
3. Project List
4. Resource update

now Project management , will just need to open the macro worksheet and run any of the following macros to get the work done !

1. any user :
-----------------
to track in how many projects he is working or has worked on ..
to which all tool he is currently having access to / or access has be revoked or not for closed project 
to track whether that employee is still working for company or left the organization


2 over all consolidated report	
-------------------------------------------
it will take the list of all the employee in map it against all the project available and populate one consolidated report with reporting capability as mentioned in step 1


3. project list
------------------
this will list all the project in the excel file and whether they are completed and all the activities completed or not ..if completed , completion date


