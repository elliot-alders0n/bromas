filename="C:\Users\Tarde\Desktop\voz-misteriosa\vocabulario\test.txt"

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(filename)

Set oShell = CreateObject("WScript.Shell")

Do Until f.AtEndOfStream
  palabra = f.ReadLine
  oShell.run "nircmd.exe mutesysvolume 0"
  oShell.run "nircmd.exe changesysvolume 100000"
  CreateObject("SAPI.SpVoice").Speak(palabra)
  wscript.sleep 500
Loop

f.Close
