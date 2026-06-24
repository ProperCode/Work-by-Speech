WScript.Sleep 1000
Set objShell = CreateObject("Wscript.Shell")
objShell.CurrentDirectory = "C:\Users\mikol\Documents\Visual Studio 2022\Projects\Work by Speech\Work by Speech\bin\Debug"
strApp = """Work by Speech.exe"""
objShell.Run(strApp)
