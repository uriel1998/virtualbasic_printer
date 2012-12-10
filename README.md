visualbasic_printer
====================

Print Return Transport Request Via Network Printer and Visual Basic Script

by Steven Saus

No warranty express or implied.

This is a VBS script to print a return transport request with user input
The printer name is currently hardcoded in.  Requires WinXP or higher,
and a VBS interpreter.  Obviously, you'll want to have permissions to print
to the printer.  Printing is handled using the system "print" command so 
that we can easily deal with non-default printers, network printers, etc.

This is the public version of the script - I have stripped all path names and
institution identifying information from it.  You will need to put in your own
printer path names and so on.  These are annotated in TOCHANGE.TXT in this Git
repository.

Please note that this script explicitly overwrites its temporary files and 
closes out objects in memory to preserve HIPPA-required privacy.

I had to research a lot of very common VBS problems to compile this script,
so I'm hoping that putting this script up will give people many examples of 
those problems.

In particular:
* Creating the equivalent of a radio button (sort of) from InputBox
* Comparing strings in VBS
* Printing to the non-default printer (including networked printers) from CLI
* Creating multi-line output in a MsgBox
* Handling case and multi-step if/then loops in VBS
* Testing for empty strings in VBS
* Testing for numeric input in VBS
* Reading and writing from text files in VBS
* Get a return status from a called system process
* Execute different actions based on the return status
* Getting human-readable system time from VBS

 Licensed under a Creative Commons BY-SA 3.0 Unported license
 To view a copy of this license, visit http://creativecommons.org/licenses/by-sa/3.0/.
