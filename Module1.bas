Attribute VB_Name = "Module1"
Option Explicit
'//****************************************//
'//  Copyright c2i - Richard CLARK
'//  http://www.c2i.fr
'//  rc@c2i.fr
'//  Code élaboré avec c2iExplorer
'//**************************************//

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
           ByVal lpClassName As String, _
           ByVal lpWindowName As String _
) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
           ByVal hwnd As Long, _
           ByVal wMsg As Long, _
           ByVal wParam As Long, _
           lParam As Any _
) As Long
Private Declare Function GetWindow Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal wCmd As Long _
) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" ( _
    ByVal hwnd As Long, _
    ByVal lpClassName As String, _
    ByVal nMaxCount As Long _
) As Long
Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_SETTEXT = &HC
Private Const WM_KEYDOWN = &H100
Private Const VK_RETURN = &HD
      
Private Const MAX_PATH = 260

Public Function GetURL() As String
Dim sIEClassName As String, hIE As Long, lngRep As Long
Dim sText As String * 255, sClass As String * 255
Dim iNum As Long, hwndChild As Long, lngRepClassName As Long
Dim lngLength As Long, sURL As String

On Error GoTo Fin
'on trouve la fenêtre de d'Internet Explorer
sIEClassName = "IEFrame"
hIE = FindWindow(sIEClassName, vbNullString)
If hIE <> 0 Then
    hwndChild = hIE
    hwndChild = hwndFindWindow(hwndChild, "WorkerW")
    If hwndChild = 0 Then Err.Raise 10
    hwndChild = hwndFindWindow(hwndChild, "ReBarWindow32")
    If hwndChild = 0 Then Err.Raise 10
    hwndChild = hwndFindWindow(hwndChild, "ComboBoxEx32")
    If hwndChild = 0 Then Err.Raise 10
    hwndChild = hwndFindWindow(hwndChild, "ComboBox")
    If hwndChild = 0 Then Err.Raise 10
    hwndChild = hwndFindWindow(hwndChild, "Edit")
    If hwndChild = 0 Then Err.Raise 10
    GetURL = ExtractURL(hwndChild)
End If
Exit Function
Fin:
MsgBox "Erreur"
End Function

Public Function SetURL(sNewURL As String)
Dim sIEClassName As String, hIE As Long, lngRep As Long
Dim sText As String * 255, sClass As String * 255
Dim iNum As Long, hwndChild As Long, lngRepClassName As Long
Dim lngLength As Long, sURL As String

On Error GoTo Fin
'on trouve la fenêtre de d'Internet Explorer
sIEClassName = "IEFrame"
hIE = FindWindow(sIEClassName, vbNullString)
If hIE <> 0 Then
    hwndChild = hIE
    hwndChild = hwndFindWindow(hwndChild, "WorkerW")
    If hwndChild = 0 Then Err.Raise 10
    hwndChild = hwndFindWindow(hwndChild, "ReBarWindow32")
    If hwndChild = 0 Then Err.Raise 10
    hwndChild = hwndFindWindow(hwndChild, "ComboBoxEx32")
    If hwndChild = 0 Then Err.Raise 10
    hwndChild = hwndFindWindow(hwndChild, "ComboBox")
    If hwndChild = 0 Then Err.Raise 10
    hwndChild = hwndFindWindow(hwndChild, "Edit")
    If hwndChild = 0 Then Err.Raise 10
    lngRep = SendMessage(hwndChild, WM_SETTEXT, 0, ByVal sNewURL)
    lngRep = SendMessage(hwndChild, WM_KEYDOWN, VK_RETURN, 0)
End If
Exit Function
Fin:
MsgBox "Erreur"
End Function

Private Function SupprimeNull(sM As String) As String
If (InStr(sM, Chr(0)) > 0) Then
   sM = Left(sM, InStr(sM, Chr(0)) - 1)
End If
SupprimeNull = sM
End Function
 
Private Function ExtractURL(hwnd As Long) As String
Dim lngLength As Long, sURL As String, lngRep As Long

lngLength = SendMessage(hwnd, WM_GETTEXTLENGTH, 0, ByVal 0)
sURL = Space(lngLength + 1)
lngRep = SendMessage(hwnd, WM_GETTEXT, lngLength + 1, ByVal sURL)
ExtractURL = SupprimeNull(sURL)
End Function

Private Function hwndFindWindow(hwndParent As Long, sClassName As String) As Long
Dim hwndChild As Long, sClass As String * MAX_PATH
Dim bTrouve As Boolean, lngRepClassName As String

'on recherche la première fenêtre enfant
hwndChild = GetWindow(hwndParent, GW_CHILD)
'on regarde la classe du premier enfant
lngRepClassName = GetClassName(hwndChild, sClass, 255)
If Left(sClass, lngRepClassName) = sClassName Then
    hwndFindWindow = hwndChild
    Exit Function
End If
If hwndChild = 0 Then Exit Function 'il n'a pas d'enfant

bTrouve = False
Do Until bTrouve
    hwndChild = GetWindow(hwndChild, GW_HWNDNEXT)
    If hwndChild = 0 Then Exit Do   'on a tout parcouru
    lngRepClassName = GetClassName(hwndChild, sClass, MAX_PATH)
    If Left(sClass, lngRepClassName) = sClassName Then
        hwndFindWindow = hwndChild  'on l'a trouvé
        Exit Function
    End If
Loop
End Function


