Attribute VB_Name = "A_FBO3"
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" _
    (ByVal hwnd As Long) As Long

Option Explicit
Public SapGuiAuto, WScript, msgcol
Public objGui  As GuiApplication
Public objConn As GuiConnection
Public session As GuiSession
Public Sub FBO3_981_mvt()

Dim SapGuiAuto As Object
    Dim objGui As Object
    Dim objConn As Object
    Dim session1 As Object
    Dim session2 As Object
    Dim wshshell As Object
    Dim matnumbers As String

matnumbers = ActiveWorkbook.ActiveSheet.Range("i1").Value

    Set SapGuiAuto = GetObject("SAPGUI")
    Set objGui = SapGuiAuto.GetScriptingEngine
    Set objConn = objGui.Children(0)
    Set session1 = objConn.Children(0)
Range("I6").Select
    Selection.Copy

session1.FindById("wnd[0]").resizeWorkingPane 133, 39, False
session1.FindById("wnd[0]/tbar[0]/okcd").text = "/nfb03"
session1.FindById("wnd[0]").sendVKey 0
session1.FindById("wnd[0]/usr/txtRF05L-BELNR").text = matnumbers
session1.FindById("wnd[0]/usr/ctxtRF05L-BUKRS").text = 5230
session1.FindById("wnd[0]").sendVKey 0
session1.FindById("wnd[0]/titl/shellcont/shell").PressContextButton "%GOS_TOOLBOX"
session1.FindById("wnd[0]/titl/shellcont/shell").SelectContextMenuItem "%GOS_VIEW_ATTA"
session1.FindById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").CurrentCellColumn = "BITM_DESCR"
session1.FindById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").SelectedRows = "0"
session1.FindById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").DoubleClickCurrentCell
Application.Wait (Now + TimeValue("0:00:06"))
'SAVING PROCEDURE

Do
    Set session2 = objConn.Children(1)
    session2.FindById("wnd[0]/shellcont/shell/shellcont[1]/shell").SetFocus
    Set wshshell = CreateObject("WScript.Shell")
    wshshell.SendKeys "^+s"
Exit Do
Loop
    Application.Wait (Now + TimeValue("0:00:02"))
    Set wshshell = CreateObject("WScript.Shell")
    wshshell.SendKeys "^v"
    session2.FindById("wnd[0]").Close
    Set wshshell = CreateObject("WScript.Shell")
    wshshell.SendKeys "{ENTER}"

End Sub
