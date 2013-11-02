excel.vim  
=========

A simple vim plugin for displaying data of an excel file.  

##NOTE:
+ It requires your vim with `python` supported.
  
+ Python module [`xlrd`](https://github.com/python-excel/xlrd) is required on your system. To install it you may run `sudo pip install xlrd`  
  
+ For `vim 7.3`, it works well for almost all kinds of file formats, ie. `.xls`,`.xlam`,`.xla`,`.xlsb`,`.xlsx`,`.xlsm`,`.xltx`,`.xltm`,`.xlt`  
  but for `vim 7.4`, only data of `.xla`, `.xls`, `.xlt` file will be parsed when opened up. For others you can run `:call ParseExcel()` in vim after opening them.  
+ It will parse all sheets by default, displaying them on different tabs.
  
+ You can only __read__ data of a file, any modifications on any tab will __NOT__ do change to the origin excel file.
  
+ Works best on excel files that contain English characters only.
