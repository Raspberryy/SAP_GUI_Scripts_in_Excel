# SAP GUI Scripts in Excel
Visual Basic Boilerplate Code to use SAP GUI Scripts in Excel. With GUI Scripting, SAP allows its Users to automate certain tasks like for example clicking through  the same menus multiple times. Running the previously recorded script can "click" these buttons for you. 

## Boilerplate Code
GUI Scripts usually require SAP to already be opened. This Boilerplate Code allows you to run an SAP GUI Script from any starting point. The script will be executed even if:

- SAP Logon is not started yet
- SAP Logon is started but the SAP Client isn't
- SAP Client is started and in home menu
- SAP Client is started but the User has something opened

Depending on the scenario, the User running the script (as VBS or in Excel) will have to accept one to two Prompts to allow the scripts access to the SAP GUI.

~~~

                                                                                        Yes -----> [Abort, don't run script]
                                                                                      /
                                                                Yes --> User wants 
                                                              /         to safe data?
                                      Yes --> Some Menu open?                         \
                                    /               ^         \                         No
              Yes --> Client Started?               |          \                         |
            /               ^       \               |           No ---> [Run Script] <----
SAP Logon                   |         No ---> Start Client
started?                    |
            \               |
              No ---> Start Logon 
~~~

## Get Started
Copy the VBS Script to wherever you will need it. In Line 8 enter the SAP Client Name. The Name can be seen in the SAP Logon. In Line 26 Confirm the installation location of the saplogon.exe. Copy your GUI Script somewhere between line 82 and 87. 

## Considerations
Whenever you are automating any task in SAP, Users might lose knowledge on how to do these tasks by themselves. Should for whatever reason the script break, a user, who previously knew the process very well, might no longer be able to debug the error and get back to the developer saying that the script isn't working anymore instead of trying to find a solution themselves.  

## Known Issues
In Line 19 the code defines SAPNotRunning: to be executed on error. This remains true for all later lines of code as well. Should there be any error in the actual GUI Script, the code returns to Line 23 (SAPNotRunning:) which eventually runs the GUI Script twice until the error happens again. Once the error happens twice, Excel/the Script will throw the actual error and abort.   
