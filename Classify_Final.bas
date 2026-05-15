Attribute VB_Name = "Module1"
Option Explicit

'========================
' Helpers
'========================
Private Function DigitsOnly(v As Variant) As String
    Dim s As String, i As Long, o As String
    s = CStr(v)
    For i = 1 To Len(s)
        If Mid$(s, i, 1) Like "#" Then o = o & Mid$(s, i, 1)
    Next i
    DigitsOnly = o
End Function

Private Function Prefix3(v As Variant) As String
    Dim t As String: t = DigitsOnly(v)
    If Len(t) >= 3 Then Prefix3 = Left$(t, 3)
End Function

Private Function Suffix7(v As Variant) As String
    Dim t As String: t = DigitsOnly(v)
    If Len(t) >= 7 Then Suffix7 = Right$(t, 7)
End Function

'========================
' MAIN
'========================
Sub Classify_Final()

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "X").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim data As Variant
    data = ws.Range("A2:AR" & lastRow).Value
    Dim n As Long: n = UBound(data, 1)

    Dim resAI() As Variant, resAJ() As Variant
    ReDim resAI(1 To n, 1 To 1)
    ReDim resAJ(1 To n, 1 To 1)

    Dim i As Long, px As Variant, sx As String

    '==================================================
    ' COUNT 986 PER SUFFIX
    '==================================================
    Dim Suffix986Cnt As Object
    Set Suffix986Cnt = CreateObject("Scripting.Dictionary")

    For i = 1 To n
        px = Prefix3(data(i, 24))
        If px = "986" Then
            sx = Suffix7(data(i, 24))
            If Not Suffix986Cnt.Exists(sx) Then Suffix986Cnt(sx) = 0
            Suffix986Cnt(sx) = Suffix986Cnt(sx) + 1
        End If
    Next i

    '==================================================
    ' ? GROUP BY CUSTOMER + ROUTE + SUFFIX (FIX)
    '==================================================
    Dim Groups As Object
    Set Groups = CreateObject("Scripting.Dictionary")

    For i = 1 To n
        Dim gKey As String
        gKey = data(i, 16) & "|" & data(i, 22) & "|" & Suffix7(data(i, 24))
        If Not Groups.Exists(gKey) Then Groups.Add gKey, CreateObject("Scripting.Dictionary")
        Groups(gKey)(i) = True
    Next i

    '========================
    ' AI LOGIC (NOW CORRECT)
    '========================
    Dim g As Variant, r As Variant

    For Each g In Groups.Keys

        Dim has986 As Boolean: has986 = False
        Dim splitPref As Object: Set splitPref = CreateObject("Scripting.Dictionary")

        For Each r In Groups(g).Keys
            px = Prefix3(data(r, 24))
            If px = "986" Then has986 = True
            If Left$(px, 1) = "8" Then splitPref(px) = True
        Next r

        Dim opt As String
        If Not has986 And splitPref.Count = 1 Then
            opt = "Option 5"
        ElseIf has986 And splitPref.Count = 0 Then
            opt = "Option 1"
        ElseIf has986 And splitPref.Count = 1 Then
            opt = "Option 2"
        ElseIf splitPref.Count >= 2 Then
            opt = "Option 4"
        Else
            opt = "Option 2"
        End If

        For Each r In Groups(g).Keys
            px = Prefix3(data(r, 24))
            sx = Suffix7(data(r, 24))

            If px = "986" And splitPref.Count = 0 And Suffix986Cnt(sx) = 1 Then
                resAI(r, 1) = "Option 1"
            Else
                resAI(r, 1) = opt
            End If
        Next r
    Next g

    '==================================================
    ' AJ - OPTION 3 / OPTION 3 SPLIT (UNCHANGED)
    '==================================================
    Dim DayGrp As Object
    Set DayGrp = CreateObject("Scripting.Dictionary")

    For i = 1 To n
        If IsDate(data(i, 9)) Then
            Dim k As String
            k = data(i, 16) & "|" & CLng(CDate(data(i, 9)))

            If Not DayGrp.Exists(k) Then
                DayGrp.Add k, CreateObject("Scripting.Dictionary")
                Set DayGrp(k)("986") = CreateObject("Scripting.Dictionary")
                DayGrp(k)("80") = False
            End If

            px = Prefix3(data(i, 24))
            sx = Suffix7(data(i, 24))

            If px = "986" Then
                DayGrp(k)("986")(sx) = i
            ElseIf Left$(px, 1) = "8" Then
                DayGrp(k)("80") = True
            End If
        End If
    Next i

    Dim dk As Variant, s As Variant

    For Each dk In DayGrp.Keys
        If DayGrp(dk)("986").Count >= 2 Then
            For Each s In DayGrp(dk)("986").Keys
                resAJ(DayGrp(dk)("986")(s), 1) = "Option 3"
            Next s
        End If

        If DayGrp(dk)("80") = True Then
            For Each s In DayGrp(dk)("986").Keys
                If Suffix986Cnt(s) = 1 Then
                    resAJ(DayGrp(dk)("986")(s), 1) = "Option 3 split"
                End If
            Next s
        End If
    Next dk

    ws.Range("AI2").Resize(n, 1).Value = resAI
    ws.Range("AJ2").Resize(n, 1).Value = resAJ

    MsgBox "? FINAL - suffix-safe classification completed", vbInformation
End Sub

