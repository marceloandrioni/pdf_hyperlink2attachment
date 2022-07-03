# pdf_hyperlink2attachment

Convert hyperlinks into attachments.


## TL;DR

__What?__ Get all the hyperlinks in a pdf file that reference local files and attach the files to the pdf.

__Why?__ This way a single pdf file can hold all the data. Also useful to attach data that does not "fit" a regular document layout (e.g.: very large tables, non-table and non-image files).

## Introduction

This simple script does the following:

* Analyzes a pdf file (input).
* Gets all the hyperlinks in the pdf file that reference local files. Remote hyperlinks (e.g.: `http:`, `https:`, `ftp:`, `mailto:`) are ignored.
* Replaces the hyperlinks with attachments.
* Saves the data changes in a new pdf file (output).

With this, all the data is stored in a single pdf file. The user can then send a single file to someone over the internet instead of sending a zip/tar package with the original pdf and the files referenced by the hyperlinks.

## Requirements
The only "real" requirements are the [pikepdf](https://pikepdf.readthedocs.io/en/latest/) and [Gooey](https://github.com/chriskiehl/Gooey) libraries. Both can be easily installed with `conda` or `pip`. All the other requirements (e.g.: `os`, `pathlib`, `warnings`, `argparse`) are part of the standard python installation.

__Note:__ `pikepdf` conda installation is only available in Linux and OSX. In Windows it should be installed with pip: `pip install pikepdf`

## How to use?

* Using Microsof Word or LibreOffice Writer, just highlight the desired text and insert an hyperlink to a local file.

* Export the `.docx` file to pdf. __Note:__ the pdf file should be exported, and not "printed" to pdf so the hyperlinks can be preserved.

* The script can then be run with in CLI or GUI mode.

__CLI__

In the command line just run: `pdf_hyperlink2attachment.py <input.pdf> <output.pdf>`

e.g.:

```
$ pdf_hyperlink2attachment.py examples/document.pdf examples/document2.pdf
Input file: examples/document.pdf
Page 1/3
  Attaching local file 'examples/tables/table2.xlsx'
  Attaching local file 'examples/tables/table1.xlsx'
Page 2/3
  Attaching local file 'examples/images/image1.png'
Page 3/3
Output file: examples/document2.pdf
Done!
```

__GUI__

Run the script with no arguments to open the GUI and then select the input and output pdf files.

<img src="./gui_example.png" alt="Firefox" width="600"/>

* Open the newly created file (e.g.: `document2.pdf`) with a pdf viewer and click on any hyperlink to access the data. The file does not depend on the existence of the local files (e.g.: `table1.xlsx`, `image1.png`) like the original pdf file (e.g.: `document.pdf`).

**_NOTE:_** Some pdf viewers (e.g.: Adobe Acrobat Reader, Firefox) show the attached files in a lateral bar, but this is not true for all viewers (e.g.: Evince). It is possible to attach files with the same name from different directories (e.g.: `dir1/myfile.txt`, `dir2/myfile.txt`). The respective hyperlinks will 
reference the correct attached files, however, the lateral attachment bar will only display the first of the homonymous files.

<img src="./firefox_example.png" alt="Firefox" width="600"/>
