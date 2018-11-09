# sample-code
just sample perl code that parses xlxs files and loads to mysql/mariadb. <br>
Initial commit written in just a couple hours but will be adding to it. 

<strong>Dependancies:</strong><br>
 DBI - perl db interface<br>
 DBD::mysql - mysql driver (should work on MariaDB as well)<br>
 SQL::Abstract - to dynamically assign placeholders; sanitize sql input without the need for static statement structures<br>
 Getopt::Long - to...get options<br>
 Spreadsheet::XLSX - for xlsx parsing and determining input data types<br> 
 Text::Iconv - converter for xlsx data<br>
 List::MoreUtils - for iterating over array elements in chunks<br>
 Data::Dump::Streamer (to debug data structures)<br><br>

<strong>Usage example:</strong>
init.pl --file F:\data_load.xlsx --database testdb --table testtable --user eparlin --pw pass --host 192.168.201.132 --port 3306<br><br>
<strong>The host and port args are optional. If you are on localhost you can omit them</strong><br><br>
<strong>TODO:</strong>
Add support/drivers for Postgres, SQLite, and newer MariaDB driver.<br>
Add usage output.<br>
Add sanitization test package.<br>
