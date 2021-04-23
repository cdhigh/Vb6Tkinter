如果打开工程TkinterDesigner.vbp失败，提示对象库未注册，一般是Windows Common Controls 6.0 (mscomctl.ocx)未成功加载。
可以先尝试注册 'regsvr32 mscomctl.ocx'，如果还不成功，则可以 'regtlib msdatsrc.tlb'。
至于 mscomctl.ocx/msdatsrc.tlb 在哪个目录，不同版本位置不同，搜索一下即可。

32bit:
cd c:\windows\system32
regtlib msdatsrc.tlb

64bit:
cd C:\Windows\SysWOW64\
regtlib msdatsrc.tlb
