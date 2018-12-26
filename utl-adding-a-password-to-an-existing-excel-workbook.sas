Adding a password to an existing excel workbook

This will work with SAS/IML/R

Output password protected work password=foo
https://tinyurl.com/y7pz2a2y
https://github.com/rogerjdeangelis/utl-adding-a-password-to-an-existing-excel-workbook/blob/master/class.xlsx

github
https://tinyurl.com/y8mcx4kv
https://github.com/rogerjdeangelis/utl-adding-a-password-to-an-existing-excel-workbook

stackOverflow SAS
https://stackoverflow.com/questions/53867436/sas-eg-export-excel-with-pass

stackOverflow SAS
https://tinyurl.com/ybf4zmbg
https://stackoverflow.com/questions/51055003/using-saveworkbook-in-r-how-do-i-add-a-password

Luke A profile
https://stackoverflow.com/users/1327739/lukea

You need to install these packages

  1. devtools::install_github("omegahat/RDCOMClient")
  2. rJava
  3. xlsx


INPUT
=====

  d:/xls/class.xlsx

      +----------------------------------------------------------------+
      |     A      |    B       |     C      |    D       |    E       |
      +----------------------------------------------------------------+
   1  | NAME       |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |
      +------------+------------+------------+------------+------------+
   2  | ALFRED     |    M       |    14      |    69      |  112.5     |
      +------------+------------+------------+------------+------------+
   3  | ALICE      |    F       |    15      |   66.5     |  112       |
      +------------+------------+------------+------------+------------+
    ...
   [CLASS]


EXANPLE OUTPUT
---------------

 A text box will pop up

A password(foo) is added to the entire workbook and
the R script opens the password protected
workbook and asks for the aforementioned scripted password(foo)
to open the workbook.


+---------------------------------------+
| D:/xls/class.xlx is protected         |
| Enter the password in you R script    |
+---------------------------------------+
|                                       |
|      +--------------------------+     |
|      | foo                      |     |
|      +--------------------------+     |
|                                       |
+---------------------------------------+


PROCESS
=======

%utl_submit_r64('
library(rJava);
library(RDCOMClient);
library(xlsx);

 set_excel_psw <- function(filename, password = rstudioapi::askForPassword()) {
  filename <- normalizePath(path.expand(filename));
  Application <- COMCreate("Excel.Application");
  wkb <- Application$Workbooks()$Open(filename);
  wkb[["Password"]] <- password;
  wkb$Save();
  Application$Quit();
  Application <- NULL;
  invisible(gc());
};
tf="d:/xls/class.xlsx";
set_excel_psw(filename = tf, password = "foo");
shell.exec(tf);
');


OUTPUT
======

same as input but with a password
see above

*                _              _       _
 _ __ ___   __ _| | _____    __| | __ _| |_ __ _
| '_ ` _ \ / _` | |/ / _ \  / _` |/ _` | __/ _` |
| | | | | | (_| |   <  __/ | (_| | (_| | || (_| |
|_| |_| |_|\__,_|_|\_\___|  \__,_|\__,_|\__\__,_|

;


libname xel "d:/xls/class.xlsx";

data xel.class;
   set sashelp.class;
run;quit;

libname xel clear;




