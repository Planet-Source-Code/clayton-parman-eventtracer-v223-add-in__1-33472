
      Project name: "EventTracer 2.0"  (An Add-In for VB 6.0)
	    Author: Clayton Parman
	      Date: 05-02-01
      Document Rev: 03-15-02  Changed wording of "Trace" commands
                              to "TrcT" and "TrcV" as per Rev 2.2

    Acknowledgment: Thanks to "Van Den Driesshe Francois" for the
		    tip on how to put an AddIn Icon on the Toolbar.


I - Description:

  "EventTracer" is basically "debug.print" on wheels. It is an Add-In
  which automatically inserts a simple command into Procedures and
  Functions, which allows you to Trace the sequence of execution via
  the "Immediate Window". Output can also be redirected to a Text file
  which allows you to follow the operation of a compiled program, (or
  compare the nuances between different sequences of events.)

  Some "optional" control and formatting parameters can be used to
  make the debug output easier to follow.

  While debugging is "EventTracer's" bread and butter, a couple of
  additional features have been included (mostly because they were
  fairly easy to do after the main body of Event Tracing code was
  completed ... hence version 2.0).  After a weekend crash course in
  Error Handling (I'm still fairly new to VB), a first go stab was
  made at adding a "Centralized Error Handler".  (See Section IV).

  The code for the Error Handler, and Event Tracing, are maintained
  in their own modules, so they can be easily upgraded, or changed
  out altogether.


II - Installation:

  1. Compile the "EventTracer.vbp" to register the .dll as an Add-In.

	Note: If you have problems getting "EventTracer" to compile,
	      make sure the "Microsoft Visual Basic Extensibility"
	      is selected in the "References" dialog box.

  2. After compiling, exit the Project "before" you run "EventTracer"

  3. Load "EventTracer" using the "Add-Ins Manager" on the "Add-Ins menu"
     and set "EventTracer's"   Load Behavior = "Loaded".

  4. If you find "EventTracer" useful, you can have set it to "Load on
     StartUp", or install it on the Add-In Toolbar (my own preference).


III - Operation:

  A. "EventTracer" makes an attempt at having a smart interface. Depending
     on conditions in the VB IDE, (and whether or not you have clicked on
     a word), "EventTracer" defaults the Focus to different buttons, at
     different Scopes. For example: If the "modEventTracer.bas" does not
     exist, then "EventTracer" defaults the focus to the "INSERT Trace for
     PROCEDURE" button, at the "Module Level". And so on.

     Bottom   Where you click "before" launching "EventTracer" does make a
      line:   difference as to which button, and what Scope it defaults to.
	      Play around with it.  If you don't like it, disable it.

  B. The "INSERT Trace for PROCEDURE" button (for Project) adds the
     "modEventTracer" module to your Project and inserts "TrcT" commands
     into each Procedure. The "modEventTracer.bas" module has Help comments
     included in it.

  C. The "INSERT Trace for procedure EXIT" button adds a 'TrcT" command to
     the end of the procedure. In effect you are creating a "call stack."
     While this can be done by Project or Module, it is most effective if
     you only add it to those procedures where you really need to follow
     the events Exit. If you are trying to trace "Lost_Focus" events, 
     you "may" need to track the Exit of most of the procedures.

  D. The "INSERT Trace for VARIABLE" works in conjunction with "TrcT",
     for printing the value of almost any variable in the Immediate Window.
     Just click on the variable name which you wish to track, then click
     the "INSERT Trace for VARIABLE" button to insert a "TrcV" command.

  E. The "INSERT Trace to File - OPEN/CLOSE" buttons insert a command
     which redirects the output from the Immediate Window to a text file
     named "Trace?.Log" (where "?" is an automatically incremented number).
     This is handy for tracing events in a compiled program (which sometimes
     trigger different events than when running from the VB IDE.)

  F. Place the OPEN command where you want to turn this feature ON (usually
     in Form_Load or Form_Initialize, but it could be anywhere. It will only
     execute one time). Place the CLOSE command where you want to turn the
     output to "Trace?.Log" OFF (usually in the procedure which has the END
     and UNLOAD FORM commands). As long as you CLOSE before OPEN, you can
     write multiple "Trace?.Log" files in a single session. This allows you
     to compare differences between various sequences of execution.

  G. Once the Event Tracing has been added to your project, you can manually
     place trace commands whereever you like just by typing "TrcT" or "TrcV".
     The "EventTracer" program is simply an automated front end for entering
     these commands.


IV - Extra Features (Independant of Event Tracing):

  A. "ADD Error Handling": The first time you add Error Handling to a
     procedure, "EventTracer" will add the module "modErrorHandler" to
     your Project, and add Error Handling to the Function(s) or
     Procedure(s). After that, Error Handling is only added to each
     Module, Function, or Procedure as you so choose.

  B. "ADD Procedure Line numbers" adds line numbers to the procedure. This
     uses the ERL Function and provides you with both the "name of the
     procedure" and the specific line of code that trigger the error.

  C. The Error Handler can also write errors to a log file. "This feature
     is turned OFF by default."  When turned ON, a new log file is auto-
     matically created for each month. Three full months of error logs are
     maintained. Logs older than that are automatically deleted.

  D. For more details about how to use, see the "modErrorHandler.bas" file
     which is created when you add Error Handling to a Project.

  E. "Close all VB Code Windows".  Does just that ... nothing more.


V - Footnote


     If you find any bugs (which you can replicate), please report them 
     to me at:

	      cdparman@mchsi.com

     I will fix it (them) and send you a fixed copy as soon as possible.

     If you have any suggestions for improvements, let me know. If there
     is any interest, I'll try to put out another upgrade.

