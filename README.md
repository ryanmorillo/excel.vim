excel.vim  
=========

A simple vim plugin for displaying data from an excel file.  

![screen shot](http://img5.tuchuang.org/uploads/2013/12/2013-12-09 21:14:57的屏幕截图.png)

## Requirements
+ It is required that vim is compiled with python3 support.
  
+ Python module [`xlrd`](https://github.com/python-excel/xlrd) is also required for old excel XLS. (This can be installed using `pip install xlrd`).

+ Python module [`pylightxl`](https://github.com/PydPiper/pylightxl) is also required. (This can be installed using `pip install pylightxl`).
  
+ For `vim 7.3` and earlier, support is good for many file formats, including:   
  `.xls`,`.xlam`,`.xla`,`.xlsx`,`.xlsm`,`.xltx`,`.xltm`,`.xlt` etc   

+ For `vim 7.4`, please add the following to your `.vimrc` file:
  ```
  let g:zipPlugin_ext = '*.zip,*.jar,*.xpi,*.ja,*.war,*.ear,*.celzip,*.oxt,*.kmz,*.wsz,*.xap,*.docx,*.docm,*.dotx,*.dotm,*.potx,*.potm,*.ppsx,*.ppsm,*.pptx,*.pptm,*.ppam,*.sldx,*.thmx,*.crtx,*.vdw,*.glox,*.gcsx,*.gqsx'
  ```

##Limitations

+ You can only __read__ data from a file, any modifications will __NOT__ be saved to the original excel file.
  
+ Works best on excel files that contain ASCII characters only.  

##Note
+ All sheets will be parsed by default, displaying them in different tabs.
