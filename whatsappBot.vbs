Dim startmundo
Set Fso = CreateObject("Scripting.FileSystemObject")
Set InputFile = fso.OpenTextFile("Bot/src/contatos.txt")
Do While Not (InputFile.atEndOfStream)
     contatos = InputFile.ReadLine
     set startmundo = CreateObject("WScript.Shell")
 wscript.sleep 5000
startmundo.SendKeys "{TAB}"
wscript.sleep 600
startmundo.SendKeys"" & contatos
wscript.sleep 3000
startmundo.SendKeys "{ENTER}"
wscript.sleep 2000
startmundo.SendKeys "^(v)"
wscript.sleep 6000  
startmundo.SendKeys "Teste de Envio"
wscript.sleep 9000
startmundo.SendKeys "{ENTER}"     
Loop
Wscript.Quit