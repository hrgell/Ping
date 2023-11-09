# Ping
Monitor if a single IP cannot be pinged

# Description
The script was copied and modified from:

https://learn.microsoft.com/en-us/previous-versions/windows/desktop/wmipicmp/win32-pingstatus

The script is written in vbscript and uses windows management instrumentation. Use the following command line to run the script:

`cscript /NOLOGO ping2.vbs`

The script tries to ping a host every 2 seconds and emits an error if the ping fails. Abort the script with Ctrl-C.
