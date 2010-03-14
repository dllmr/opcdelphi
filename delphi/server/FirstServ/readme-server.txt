18 January 2008

It has been a few years since my last look at this code and I thought it time to update the source for Delphi 2007 SP3. I also made some changes for additional testing performed.

I can be reached at zayin@pdq.net.

Enjoy,

Mark

Everest Software LLC
http://www.hmisys.com

While every precaution has been taken in the preparation of this example/source code, Everest Software LLC and its employees assume no responsibility for errors or omissions.  Neither is any liability assumed for damages resulting from the use of the information contained herein.


6 September, 2002

Small changes to allow Delphi 7.0 build.


5 July, 2001

Now compatible with Delphi 6.0 and 5.01.


14 March, 2001

I have run the server through the compliance test tool.  I made the needed changes to pass the
test.  It now passes with just the warning(s) in data types.  I have been told the test will be
changed in the next version of the test tool.  I used version 1.0.1070 of the Compliance Test
Tool.

One of the changes was to revert back to "flat" data storage.  I have done a real hierarchical
system and I did not have the time to implement it in this example.  I also added the calls to
prevent an error when terminating the example while clients are connected.  I also incorporated
the new headers Mike created.

Enjoy.


31 January, 2000

Changed the browse method to "OPC_NS_HIERARCHIAL" which led to some other changes in data
storage and collection.

Added support for multiple clients.


7 October, 1999

I wrote this server code when I was learning COM and OPC.  Errors in the code may be present and
my interpretation/implementation of the specification for OPC Data Access 2.03, July 27, 1999
may also contain FUBARs.  If you find any discrepancies or plain old errors please let me know
what you found.

This code is released with the purpose:  give others a working example of how to write an OPC
server using Delphi and OOP.  It is intended to be an aid in learning how to write a server.
It is not intended to be a template or a base for a product.  If you wish to include any of this
code in a product please contact Fire and Safety International, Systems, Inc.
http://www.fsi-corp.com

OOP is a great language and Delphi is a wonderful product for implementing OPC.

Several people helped with an example or just an ear and a comment.  I have never shaken hands
with any of these guys.  I say to them "Thank you".

D.G. Somerton
Astrit Shapiri
John Romedahl
Nickolas Robinson
Roland Miezianko
Mike Dillamore

I tested the server with a program named Visual OPCTest Client.  It is a great program.
Roland Miezianko is the author.  http://www.opctest.com is the web address.

I can be reached at zayin@pdq.net.

Enjoy the code,

Mark R. Drake
Fire and Safety International, Systems, Inc.


Some other data:

NT4 SP5 or later
Delphi 5.01, 6.02 or Delphi 2007 SP3
You will need the OPC supplied DLLs:  http://www.opcfoundation.org/
You will also need the Delphi conversions of the OPC interfaces:
http://www.opcconnect.com/delphi.php

While every precaution has been taken in the preparation of this example/source code,
Fire and Safety International, Systems, Inc. and its employees assume no responsibility for
errors or omissions.  Neither is any liability assumed for damages resulting from the use of the
information contained herein.