
<html><center><h1>Tools for mass mail creation in one click.</h1>
<br>
<h>Almost done. </h>
<br>
<h2>GUI - done</h2> 
<h2>INI - done</h2>
<h2>Reading input fro file- done</h2>
<h2>Logging-done</h2>
<h2>Templates filtering - done</h2>
<h2>Sending to outlook-done</h2>
</center>
<b style="font-size: 10; vertical-align: auto;">
Hello!<br>
This is tool for users who need sens alot of emails with constants types for one event.
<br>
For example:
<br>
You need to send 10 emails to different pipl\company to start\stop same process.
<br>
But you need to use current data\name of responsiblity person\tecknical details\.. for this in all emails.
<br>
This tool for you!
You create templates,discribe a fields once, and create emails with actual data in them.

<br>
<br>
<h3>omc.ini - main configuration file</h3><br>
[files] - described used files <br>
fields = fields.ini - where fields list is<br>
ico = omc.ico - ico file to use in main window<br>
log = omc.log - where logs will be stored<br>
inputdata=input.ini - file with data for automaticly load<br>
emailtemlates=emailtemlates.ini - where email templates list<br>
[params] - only one param - verbose<br>
verbose=1 - if enabled(1) more info will be stored in logfile<br>

<h3>fields.ini - file with fields description.<br>
This file used for generating Gui input forms and for filtering templates.</h3><br>
For example:<br>
Field < tab > describtion<br>
StoreNum	Store SAP number<br>
StorePhone	Store phone number<br>
StoreAddress	Store full address<br>

<h3>inputdata.ini - holds all real data for all fields, described in fields.ini, not required you can input data manualy in GUI field.</h3><br>
For example:<br>
Field < tab > data<br>
StoreNum	2072<br>
StoreName	Some name<br>
StoreAddress	Гдето в городе<br>
StoreCity	Копенгаген<br>

<h3>emailtemlates.ini - holds path to email templates to use.</h3>
<br>
You can use relative path.<br>
Consist from template name < tab > file path<br>
Example:<br>
operators	tmpl\operators.emtpl<br>
Where "operators" is template name and "tmpl\operators.emtpl" file path to it.<br>

<h3>Template file</h3>
This file contains 3 fileds.<br>
First line is TO:<br>
Second line is CC: <br>
Third line is Subject<br>
All other lines is email body in html format.<br>
Example:<br>
Some person someperson@company.net<br>
Some person 2 someperson2@company.net<br>
We are closing StoreName.<br>
Hello please do the needful to start process on StoreName StoreNum<br>
Regards.
</b>

</html>
