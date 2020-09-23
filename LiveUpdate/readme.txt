'****************************************************************
'
' Live Update Code
'
' Written by:  Francisco H Tapia
'              3/11/2003
' Special Thanks to Blake Pell <blakepell@hotmail.com> for inspiring this
' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=13413&lngWId=1
' This code is open source, I would appreciate that anybody using
' this is a released application to e-mail or get in contact with
' me.
' The original concept of this code and much of the Form code is very similar to that of Blake's
' I wanted to use an INI FILE source for the URL paths of the web and destination files because
' I did not want to keep re-compiling the project for other programs...
' Also the space where Version.ver downloads into is sperate from the program so that if there
' are errors the program will re-download the update instead of staying broken.
' This version of the Live Update code, does away with user interaction, except to ask the
' user if they want to continue w/ the D/L.
' I hope this makes someone's day easier or helps them learn
' a bit as it did for me.
' WININET.DLL SOURCE:  http://support.microsoft.com/default.aspx?scid=KB;EN-US;Q232194&
'
'
'****************************************************************

So what's this code do?
If you have a file server and not a webserver this is pretty handy... another thing also is that it executes the file you just downloaded... I'm currently working on building this with an asynchronous process to help update the form because it will appear as if it is frozen but that's because of the wininet.dll download...

VERY IMPORTANT, check the INI file and set the correct path to your server and URL paths to your web servers if you will be using them, otherwise ignore those fields.  But you MUST set the correct path to your Local file destinations, this is where you will place the files.

Version 1.1
	Fixed If If IsConnected = false Then error
	added a kill previous d/l copy of file before actual download to avoid mismatch errors

Version 1.0
	posted to PSC