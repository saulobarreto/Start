Set WshShell = WScript.CreateObject("WScript.Shell")

Dim exeName1
Dim statusCode1
Dim exeName2
Dim statusCode2
Dim exeName3
Dim statusCode3
Dim exeName4
Dim statusCode4
Dim exeName5
Dim statusCode5

exeName1 = """C:\Program Files\Internet Explorer\iexplore.exe"""
exeName2 = """C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"""
exeName3 = """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
exeName4 = """C:\Users\saulo.b\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Spyder.lnk"""
exeName5 = """C:\Users\saulo.b\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\TEams.lnk"""

statusCode1 = WshShell.Run (exeName1, 1, false)
statusCode2 = WshShell.Run (exeName2, 1, false)
statusCode3 = WshShell.Run (exeName3, 1, false)
statusCode4 = WshShell.Run (exeName4, 1, false)
statusCode5 = WshShell.Run (exeName5, 1, false)

'MsgBox("Done!")