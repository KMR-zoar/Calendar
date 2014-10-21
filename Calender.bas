Attribute VB_Name = "Calender"
' Zoar.

Sub Calender()
   Dim Obj As Object
   Dim corner(0 To 2) As Double
   Dim LtDay As Long
   Dim DayofWeek As Long
   Dim i As Long
   Dim charSet As Long
   Dim PFamily As Long
   Dim Caltext As String
   Dim strDay As String
   Dim nameFont As String
   Dim Bold As Boolean
   Dim Italic As Boolean
   
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
   ThisDrawing.ActiveTextStyle.SetFont "ÇlÇr ÉSÉVÉbÉN", Bold, Italic, charSet, PFamily
   
   Set Obj = ThisDrawing.ModelSpace.AddMText(corner, 35, Caltext)
   
   ZoomExtents
End Sub

