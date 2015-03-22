Option Explicit
Dim Excel, GetMessagePos, x, y, Count, Position

Do While Count < 10
 Set Excel = Wscript.CreateObject("Excel.Application")

 GetMessagePos = excel.ExecuteExcel4Macro("CALL(""user32"",""GetMessagePos"",""J"")")

 x = CLng("&H" & Right(Hex(GetMessagePos), 4))
 y = CLng("&H" & Left(Hex(GetMessagePos), (Len(Hex(GetMessagePos)) - 4)))
 If Count MOD 2 = 0 Then
  Position = "- 30"
 Else
  Position = "+ 30"
 End If
 Excel.ExecuteExcel4Macro("CALL(""user32"",""SetCursorPos"",""JJJ""," & x & " " & Position & "," & y & " " & Position & ")")

 WScript.Sleep(1000)
 Count = Count + 1
Loop

WScript.Echo "Program Ended"
