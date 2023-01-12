filename="C:\tetas.txt"

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(filename)

Do Until f.AtEndOfStream
  palabra = f.ReadLine
  CreateObject("SAPI.SpVoice").Speak(palabra)
  wscript.sleep 500
Loop

f.Close
