                            EventTracer rev 2.2            04-02-2002


A note from the author: 

          While any bug is a nuisance, the nature of the bugs that usually 
          appear in EventTracer seriously detract from its good points.
          While every effort has been made not to turn PSC supporters into
          beta testers, the huge combinations of code and programming styles
          that EventTracer processes have triggered a lot of bugs. If you 
          should happen to encounter a bug, please report it to me at 
          "cdparman@mchsi.com" and I will make every effort to fix it ASAP
          and upload a patch to PSC.

          Clayton Parman



Bug fixes/modifications since 2.2 release:


          version 2.2.3

05-26-02: Found and fixed a bug created by me, which caused an endless loop
          when trying to delete EventTracer commands. This bug was created in 
          the version 2.2 release (when trying to speed up code processing.)

05-26-02: Fixed bug in "InsertLineNumbers" procedure. During random testing,
          on some source code downloaded from PSC, ran into a an empty
          sub procedure in the code, left behind by the programmer. Adding 
          line numbers to this procedure triggered a runtime error 5.

05-26-02: Fixed bug in "GetFirstLastLinesOf" procedure. During random testing, 
          on some source code downloaded from PSC, discovered that the VBIDE 
          would return two different line counts for dead identical pieces of 
          code (in two different modules). 

05-26-02: Added a custom version of the EventTracer commands "back" to the 
          EventTracer module. EventTracing was removed when the first release
          of EventTracer was finished, but will be left in permanently now.

05-25-02: Modified EventTracer to allow adding Line numbers to procedures
          that did not have an ErrorHandling routine. (As per suggestion
          by Ian Brooke).

05-25-02: In the process of testing the above modification, ran into an
          instance of the VBIDE incorrectly reporting the Starting Line#
          of a procedure. Added a fix/workaround for this bug in the VBIDE.

05-25-02: Trying to add a "Trace for Procedure" to a "Property statement"
          would trigger a series of errors. For now, and the foreseeable
          future, EventTracer WILL NOT be working with "Property state-
          ments". The bug has been fixed to not trigger an error. 
          (Bug reported by Ian Brooke).

05-25-02: When adding "Trace for Procedure" to a "Sub" with no optionals
          in front of it (Public, Private, Friend or Static) you would
          get a runtime error 5.  This resulted from the "alignment"
          of TrcT commands directly under the Sub name (introduced in
          Revision 2.2)  (Bug reported by Ian Brooke).

05-25-02: Was incorrectly adding error line numbering to #if-#endif
          constructs. (Bug reported by Ian Brooke)


          version 2.2.2

05-12-02: When Adding EventTracing to a brand new project: Clicking
          on the "Add Trace for Procedure" at the "Procedure level" 
          (the default scope) would trigger a runtime error 5. Changing
          the default scope "module level" for new projects provided
          an simple and effective fix.


          version 2.2.1

04-05-02: If the stars were aligned just right, some logic for the
          interface would try to set the focus on a button that had
          been disabled. This triggered the beloved runtime error 5.
          Fixed the problem by simply testing to see if the button was
          enabled "before" trying to set the focus.



Improvements:

          version 2.2.0   04-02-02

1. Made some minor modifications to the "Readme v20.txt" documentation to 
   reflect the wording changes outlined in the next paragraph.

   NOTE: If you are new to EventTracer, you should still refer to the 2.0
         documentation for installation and operational reference.

2. Renamed all occurances of "Trace" to "TrcT" and all "TraceVar" to "TrcV".
   "TrcT" is abbreviation for "Trace Text". It is essentially a text message
   which prints to the Immediate window. It is used for Procedures, Functions,
   and any "manual comments" needed for following the execution of your code.

   If you are currently using EventTracer, you should do a global replace 
   for all occurances of "Trace" and "TraceVar" ... and delete the current 
   "modEventTracer.bas" module and let EventTracer 2.2 create a new one.
   (After deleting the module, just Insert a Trace for procedure or variable
    and EventTracer will automatically create a new "modEventTracer.bas").
   
   NOTE: This eliminates any potential conflicts with code created by other
         programmers which contain the word "Trace". It is not nearly as 
         likely that someone else would use the names "TrcT" or "TrcV".


3. Changed the default "scope" from "Project level" to "Procedure level" or
   "Module level". I've been using EventTracer to debug a large project and
   this seems to be the scope I want to use the most often. I haven't found
   any use for performing operations at the "Project level" and to the con-
   trary, find it's best avoided.

4. Added an option to insert the Trace for a Variable "before" the current
   line of source code instead of having insert "after" as the only option.

5. Changed the insertion points for "TRACE Procedure" commands to align 
   directly under the procedure names (to make the source code more readable).

6. Added a few lines of code which always returns you to the current line of 
   source code you were working on before having invoked a Module or Project
   level operation. This makes EventTracer's operation a bit more transparent.

7. Changes in the "Trace for Procedure [Exit]" command:

   a. For exiting Functions: Now includes the "return value" of the function.

   b. If you click on the word "Exit" (for either "Exit Sub" or "Exit Func-
      tion"), before launching EventTracer, it will default to "Insert 
      Trace Above Variable". HOWEVER, instead of inserting a Trace for a 
      Variable it will insert a "Trace for Procedure [Exit]" command. 

      Note: This is an alternative for those offended by section 7c below.

   c. When adding a Trace for Exit, if the procedure contains an "Exit Sub" 
      or "Exit Function" command, then a label is added to the end of the
      procedure and the "Exit" commands are changed to "GoTo ExitLabel"
      commands (these can be easily converted back to their original form
      with "Search and Replace").
	
      A note to the "NO GOTO!" zealots:

      I originally created the method in step 7b... but didn't like it. In
      my opinion, using a "GoTo" produces much cleaner code. It also works 
      better if a recordset needs to be closed, or the Screen.MousePointer 
      needs to be set back to "default" before exiting a procedure. If a
      "GoTo" is used, these types of things can be handled just once, at 
      the end of the procedure.
     


Bugs fixed:

1. The loop(s) for turning Trace commands ON and OFF for a Module or Project
   were very inefficient. For small modules, the problem was not evident.
   However, large modules would bring the computer to its knees. Very large 
   modules "appeared" to hang the computer up.

   Note: Turning Traces ON and OFF could be made even faster yet, but it
         would be at the expense of making the memory footprint a bit bigger.

2. The "modErrorHandler.bas" (which is created by EventTracer when you add
   Error handling to your application) had a typo in it for the handling  of
   "ObjectErrors", causing the error handler to respond to them as 
   "PrintErrors". If you are using this module, you can delete it and let
   EventTracer create a new one for you (this happens whenever you add error
   handling to a procedure).

3. Fixed a couple of bugs in the "intelligent defaults". EventTracer was not
   always producing the automatic default that I expected it to. Seems to be
   working pretty well now, (even though it doesn't always select the default
   I think it should).

4. EventTracer was not executing the "Insert Trace for Procedure" routine
   if a procedure already had a "Trace for Procedure EXIT" command.

5. Under just the right conditions, EventTracer would hang in an endless
   loop when "Turning Trace commands OFF." For me, this was about a 1 in 500
   occurance. I'm not clear why it didn't happen all the time, instead of
   just sometimes. None-the-less, I think it's fixed now.

6. When adding "Error Line Numbers" to a procedure with a "Select Case"
   block, Visual Basic does not allow the first "Case" statment to have a
   line number in front of it. Fixed EventTracer to "not" add a line number 
   to the first "Case" statement in the block. (Also added code to auto-
   matically indent the first "Case" statement to keep it consistant with 
   the rest of the code formatting).



Comments:


   I have been putting EventTracer through some fairly heavy usage. On two 
   occassions EventTracer crashed during use. It "seems" to have had a con-
   flict with the VB IDE. No harm was done to the source code I was working
   on and I was unable to reproduce either crash. Crashes have been very 
   rare, non-destructive, and are probably not fixable (certainly not worth
   the amount of effort it would require to go after them.)

   I ran into an ADO "duplicate record error" which the Error Handling
   module would not trap. I also tried the "HuntERR30" error module (also 
   available on Planet-Source-Code) and it would not trap it either. I
   fixed my application to just not produce the duplicate record and have 
   not pursued finding out why this particular error could not be trapped.


