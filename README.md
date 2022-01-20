# pingandlog

Simple vbscript to ping some hosts and log only errors to txt

Download script and execute in command line:

**cscript.exe pingandlog.vbs**


The script ping some hard coded hosts to check network connectivity.
Only failures are logged to file to reduce unuseful info.

Script use native Win32_PingStatus class from Microsoft not using external commands.

Reference:
https://docs.microsoft.com/en-us/previous-versions/windows/desktop/wmipicmp/win32-pingstatus
