Attribute VB_Name = "A_DOWNLOADALL"
Public Sub ForAllMvt_Click()
    Dim counter As Long
    Dim maxiterations As Long

    maxiterations = ActiveWorkbook.ActiveSheet.Range("J4").Value
' lOOPING PROCEDURE
    For counter = 1 To maxiterations
        If IsNumeric(Range("J1").Value) Then
            Range("J1").Value = Range("J1").Value + 1
        ElseIf Range("J1").Value = "No" Then
            Range("J1").Value = 1
        End If

        If IsNumeric(Range("J1").Value) And IsNumeric(Range("J4").Value) Then
            If Range("J1").Value > Range("J4").Value Then
                MsgBox "Completed"
                Exit Sub
            End If
        End If
'STOP IF N/A OR NO LONGER HAVE INVOIES

        Do While IsError(Range("I1").Value)
            If IsNumeric(Range("J1").Value) Then
                Range("J1").Value = Range("J1").Value + 1
            End If
            Application.Wait (Now + TimeValue("0:00:01"))

            If IsNumeric(Range("J1").Value) And IsNumeric(Range("J4").Value) Then
                If Range("J1").Value > Range("J4").Value Then
                    MsgBox "Completed"
                End If
            End If
        Loop
'DECIDE EITHER 981 MVT(USE FB03) OR 101 MVT(USE VF03)
        Dim mvtType As String
        mvtType = ActiveWorkbook.ActiveSheet.Range("I3").Value

        If mvtType = "101" Then
            A_VFO3.VFO3_101_mvt
        ElseIf mvtType = "981" Then
            A_FBO3.FBO3_981_mvt
        End If
    Next counter
End Sub
