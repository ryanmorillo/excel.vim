" Author: [yakiang](http://yakiang.com)
" Contact : strak47@gmail.com


if !has('python')
    echo 'excel.vim requires support for python'
    finish
endif

au BufRead,BufNewFile *.xls exec "%d" | set nowrap

python << EOF
import vim, xrld
EOF


echo "excel"








function! YoudaoFanyi()
    let word = eval("expand('<cword>')")
python << EOF
#-*-coding:utf-8-*- 
import vim, urllib2, json
url = 'http://fanyi.youdao.com/openapi.do?keyfrom=yakiang&key=1402135116&type=data&doctype=json&version=1.1&q=%s'%vim.eval('word')
errorCode = {20:'text too long', 30:'something wrong', 40:'unsupported language',\
                50:'Youdao key error. please update this plugin'}
try:
    data = urllib2.urlopen(url).read().decode('utf-8')
    data = json.loads(data)
    if data['errorCode'] in errorCode.keys():
        print errorCode[data['errorCode']]
    elif not data.has_key('basic'):
        print 'not found'
    else:
        meaning = data['basic']['explains']
        para = '\n'.join(meaning).encode('utf-8')
        print para
except Exception, e:
    print str(e)
EOF
endfunction

