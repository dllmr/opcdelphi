
Welcome to sOPC
---------------

sOPC is a prototype OPC Data Access Server 2.0 Toolkit for Delphi.
sOPC has been tested up to Delphi 2010.


Content
-------

Source  - sOPC Toolkit source code
Example - sOPC demo program

You will also need Delphi conversions of the OPC Custom Interfaces
from OPC Programmers' Connection

(http://www.opcconnect.com/delphi.php) and the OPC proxy/stub DLLs
from the OPC Foundation (http://www.opcfoundation.org/).


Notes
-----

When creating your own OPC server, you must Define new 'GUID' numbers
in the type library editor for 'IOPCDataAccess20' and
'OPCDataAccess20'.  This is done using the right mouse button and the
'New GUID' function.  While using the type library editor, you should
also change the Help Strings as appropriate for your application.

The OPC server object is initialized in the main project source file
(in the demo project this file is called DataAccess20Demo.dpr).  You
should choose your own unique server name and description when
creating your OPC server.

Functions to define the server's name space, and to perform read and
write operations, are found in the sample unit 'uOPCDemo.pas'.  You
will need to replace these with the code required by your own
application.

Finally, don't forget to register your OPC server before testing it:
execute your program with the command line parameter '/regserver'.


Gerhard Schmid

mailto:Gerhard.Schmid@Schmid-ITM.de
