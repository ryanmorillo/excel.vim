excel.vim  
=========

A simple plugin for reading data of an excel file.  

##NOTE:
+ It requires your vim with `python` supported.
  
+ Python module [`xlrd`](https://github.com/python-excel/xlrd) is required on your system. To install it you may run `sudo pip install xlrd`  
  
+ When open up an excel file (`.xls` etc), it will parse all sheets by default, displaying them on different tabs.
  
+ You can only __read__ data of a file, any modifications on any tab will __NOT__ do change to the origin excel file.
  
+ Works best on excel files that contain English characters only.
