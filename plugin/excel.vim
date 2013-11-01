" Author: [yakiang](http://yakiang.com)
" Contact : strak47@gmail.com

if !has("python")
    echo "excel.vim requires support for python"
    finish
endif

au BufRead,BufNewFile *.xls :call ParseExcel()


function! ParseExcel()
set nowrap
:python import xlrd
python << EOF
import vim
#-*-coding:utf-8-*- 

def getRealLengh(str):
    length = len(str)
    for s in str:
        if ord(s) > 256:
            length += 1
    return length

vim.command("let currfile = expand('%:p')")
currfile = vim.eval("currfile")
excelobj = xlrd.open_workbook(currfile)
for sheet in excelobj.sheet_names():
    cmd = "tabedit %s" % (sheet)
    vim.command(cmd)
    shn = excelobj.sheet_by_name(sheet)
    rowsnum = shn.nrows
    for n in xrange(rowsnum):
        line = ""
        for val in shn.row_values(n):
            val = isinstance(val,  basestring) and val.strip() or str(val).strip()
            line += val + ' ' * (30 - getRealLengh(val))
        vim.command("let r = append(line('$'),  '%s')"%line)

for i in xrange(excelobj.nsheets):
    vim.command("tabp")
vim.command("q!")

EOF
endfunction
