Set auto=CreateObject("wscript.shell")
CreateObject("WScript.Shell").Run WScript.ScriptFullName
CreateObject("WScript.Shell").Run WScript.ScriptFullName
CreateObject("WScript.Shell").Run WScript.ScriptFullName
CreateObject("WScript.Shell").Run WScript.ScriptFullName
Dim Excel: Set Excel = WScript.CreateObject("Excel.Application") 
do
Excel.ExecuteExcel4Macro "CALL(""user32"",""SetCursorPos"",""JJJ"",""0"",""0"")"
auto.sendkeys "^%l"
Excel.ExecuteExcel4Macro "CALL(""user32"",""SetCursorPos"",""JJJ"",""5000"",""0"")"
auto.sendkeys "^%l"
Excel.ExecuteExcel4Macro "CALL(""user32"",""SetCursorPos"",""JJJ"",""5000"",""5000"")"
auto.sendkeys "^%l"
Excel.ExecuteExcel4Macro "CALL(""user32"",""SetCursorPos"",""JJJ"",""0"",""5000"")"
auto.sendkeys "^%l"
loop
