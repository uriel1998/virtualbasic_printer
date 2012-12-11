#Return Transport Request Documentation (1)

##Name

return_request.vbs - Get user input to create a return transport request and print to a predefined Windows printer.

##Synopsis

**return_request.vbs**

The script has no command-line arguments.

##Description

*return_request.vbs* is a Visual Basic script that gets user input to create a specific form and print that form to a predefined Windows printer.&nbsp; It obtains user input from a series of GUI input boxes.&nbsp; The script returns a messagebox with a success or error condition for the user. 

This script is indicated for requesting transport originating from a remote holding location or department to the patient's room or other location.&nbsp; This script is not intended for use when the patient transport originates from the patient's room.&nbsp; 

Existing processes do not have the functionality to request return a patient from a remote holding room to their room.&nbsp; Telephonic notification creates inefficiencies in workflows for multiple departments.

Exiting the script is achieved by leaving any input box blank.

The public version of the script and documentation is available at [https://github.com/uriel1998/virtualbasic_printer](https://github.com/uriel1998/virtualbasic_printer)

##HIPPA Concerns

Several methods are utilized to prevent leakage of HIPPA data:

+ Printer locations hardcoded into the script itself ; end-users cannot alter the printing locations.
+ Temporary files are written immediately prior to printing and explicitly deleted afterward.
+ Memory allocations are expressly released and assigned to null values.
+ No logfile is created or maintained.

##Prerequisites

Several prerequisites are necessary to utilize this script properly:

+ Microsoft Visual Basic 6 (c:\WINDOWS\system32\msvbm60.dll)
+ User permissions to print to desired printer(s)
+ Temporary directory located at c:\temp

Visual Basic 6 is installed on many Windows platforms by default, and has been since 1998.&nbsp; Some newer versions of Windows may not have VB6 installed;&nbsp; one guide to installing it on Windows7 is located at:&nbsp; [http://www.fortypoundhead.com/showcontent.asp?artid=20502](http://www.fortypoundhead.com/showcontent.asp?artid=20502)

The copy of the script on GitHub has all institutional identifying data removed.&nbsp; A number of small customizations will need to be made to the script in order to have it operate correctly on your system.&nbsp; The list of these modifications (with line numbers) is in the TOCHANGE.TXT file in the GitHub repository.&nbsp; The public version of the script is also named return_print_public.vbs instead of return_print.vbs .

##License

This script and all documentation is licensed under a Creative Commons BY-SA 3.0 Unported license.&nbsp; 
To view a copy of this license, visit [http://creativecommons.org/licenses/by-sa/3.0/](http://creativecommons.org/licenses/by-sa/3.0/).

</html>