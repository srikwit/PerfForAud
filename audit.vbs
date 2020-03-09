'Reference:https://blogs.technet.microsoft.com/askperf/2012/02/17/useful-wmic-queries/

Set WshShell = WScript.CreateObject("WScript.Shell")
WScript.Echo ""
WScript.Echo " ____ ____ ____ ____ ____ ____ ____ "
WScript.Echo "||s |||r |||i |||k |||w |||i |||t ||"
WScript.Echo "||__|||__|||__|||__|||__|||__|||__||"
WScript.Echo "|/__\|/__\|/__\|/__\|/__\|/__\|/__\|"
WScript.Echo ""

WScript.Echo "[+] Preparing directories for storage"
return = WshShell.Run("cmd /c mkdir list",0,true)
return = WshShell.Run("cmd /c mkdir csv",0,true)


WScript.Echo "[+] Initializing queries"
Dim queries
'Add datafile,fsdir,ntevent,server to the query list if time permits
queries = Array("baseboard","bios","bootconfig","cdrom","computersystem","cpu","dcomapp","desktop","desktopmonitor","diskdrive","diskquota","environment","group","idecontroller","job","loadorder","logicaldisk","memcache","memlogical","memphysical","netclient","netlogin","netprotocol","netuse","nic","nicconfig","ntdomain","onboarddevice","os","pagefile","pagefileset","partition","printer","printjob","process","product","qfe","quotastring","recoveros","Registry","scsicontroller","service","share","sounddev","startup","sysaccount","sysdriver","systemenclosure","systemslot","tapedrive","timezone","useraccount","memoryaccount")

WScript.Echo ""

For Each query in queries
	WScript.Echo "[+] Fetching "&query
	return = WshShell.Run("cmd /c wmic "&query&" list full /format:list > %cd%/list/"&query&".txt ",0,true)
	return = WshShell.Run("cmd /c wmic "&query&" list full /format:csv > %cd%/csv/"&query&".csv ",0,true)
Next

WScript.Echo ""
WScript.Echo "[+] Done!"
