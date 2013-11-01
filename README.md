excel.vim  
=========

A simple plugin for reading data of an excel file.  

##NOTE:
+ It requires your vim with `python` supported.
+ Python module [`xlrd`](https://github.com/python-excel/xlrd) is required in your system. To install it you may run `sudo pip install xlrd`  
+ When open up an excel file (`.xls` etc), it will parse all sheets by default, displaying them on different tabs.
+ You can only __read__ data in an file, any modifications on any tab won't do change to the origin excel file.
+ Works best on excel files that contain English characters only.
