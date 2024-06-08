Attribute VB_Name = "A_VFO3"
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" _
    (ByVal hwnd As Long) As Long

Option Explicit
Public SapGuiAuto, WScript, msgcol
Public objGui  As GuiApplication
Public objConn As GuiConnection
Public session As GuiSession

Public Sub VFO3_101_mvt()
'VF03 INVOICE GETTER - 101 MVT/INTERCOMPANY
    
Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)

Dim AcrobatWindowName As String
Dim saveWindowName As String
Dim PrintWindowName As String
    Dim printHwnd As Long
    Dim saveHwnd As Long
Dim refNumbers
Dim wshshell
Dim hwnd As Long
    Dim windowTitle As String
    Dim startTime As Double

refNumbers = ActiveWorkbook.ActiveSheet.Range("i5").Value
 Range("I6").Select
    Selection.Copy

session.FindById("wnd[0]").resizeWorkingPane 133, 39, False
session.FindById("wnd[0]/tbar[0]/okcd").text = "/nvf03"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/usr/ctxtVBRK-VBELN").text = refNumbers
session.FindById("wnd[0]/usr/ctxtVBRK-VBELN").caretPosition = 9
session.FindById("wnd[0]/mbar/menu[0]/menu[11]").Select
session.FindById("wnd[1]/usr/tblSAPLVMSGTABCONTROL").GetAbsoluteRow(0).Selected = True
session.FindById("wnd[1]/tbar[0]/btn[6]").Press
session.FindById("wnd[2]/usr/ctxtNAST-LDEST").text = "locl"
session.FindById("wnd[2]/usr/chkNAST-DIMME").Selected = vbTrue
session.FindById("wnd[2]/usr/chkNAST-DELET").Selected = vbTrue
session.FindById("wnd[2]/usr/ctxtNAST-LDEST").text = "locl"
session.FindById("wnd[2]/usr/chkNAST-DELET").SetFocus
session.FindById("wnd[2]/tbar[0]/btn[0]").Press

session.FindById("wnd[1]/tbar[0]/btn[86]").Press

'SAVING PROCEDURE
windowTitle = "Print"

    Do

        hwnd = FindWindow(vbNullString, windowTitle)
        
        If hwnd <> 0 Then
            If SetForegroundWindow(hwnd) <> 0 Then
                Set wshshell = CreateObject("WScript.Shell")
                Application.Wait (Now + TimeValue("0:00:01"))
                Set wshshell = CreateObject("WScript.Shell")
                wshshell.SendKeys "{ENTER}"
                Application.Wait (Now + TimeValue("0:00:02"))
                Set wshshell = CreateObject("WScript.Shell")
                wshshell.SendKeys "^v"
                Application.Wait (Now + TimeValue("0:00:01"))
                Set wshshell = CreateObject("WScript.Shell")
                wshshell.SendKeys "{ENTER}"
                Exit Do
            Else
                MsgBox "Failed to bring Print window to the front."
                Exit Do
            End If
        End If
        Application.Wait (Now + TimeValue("0:00:02"))
    Loop
End Sub
