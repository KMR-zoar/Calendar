Attribute VB_Name = "Calender"
' Zoar.

Sub Calender()
   Dim Obj As Object, corner(0 To 2) As Double, i As Long
   Dim LtDay As Long, DayofWeek As Long, charSet As Long
   Dim PFamily As Long, Caltext As String, strDay As String
   Dim nameFont As String, Bold As Boolean, Italic As Boolean
   
   LtDay = Day(DateSerial(Year(Date), Month(Date) + 1, 1) - 1)
   
   DayofWeek = Weekday(DateSerial(Year(Date), Month(Date), i))
   Caltext = String(3 * DayofWeek, " ")
   
   For i = 1 To LtDay
      DayofWeek = Weekday(DateSerial(Year(Date), Month(Date), i))
      
      Caltext = Caltext & Right("  " & CStr(i), 3)
      
      If (DayofWeek Mod 7) = 0 Then
         Caltext = Caltext & vbCrLf
      End If
   Next i
   
   corner(0) = 0: corner(1) = 20: corner(2) = 0
   
   ThisDrawing.ActiveTextStyle.GetFont nameFont, Bold, Italic, charSet, PFamily
   ThisDrawing.ActiveTextStyle.SetFont "�l�r �S�V�b�N", Bold, Italic, charSet, PFamily
   
   Set Obj = ThisDrawing.ModelSpace.AddMText(corner, 35, Caltext)
   
   ZoomExtents
End Sub

