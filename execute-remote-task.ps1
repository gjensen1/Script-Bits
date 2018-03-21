schtasks /create /s $ServerName /tn "bsa enable psremoting" /tr "powershell enable-psremoting -force" /sc once /st 19:00 /ru SYSTEM

schtasks /run /s $ServerName /tn "bsa enable psremoting"

#powershell -command "& {start-sleep -s 60}"

schtasks /delete /s $ServerName /tn "bsa enable psremoting" /f