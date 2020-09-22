Attribute VB_Name = "codemodule"
Option Explicit
'
' Module Declares for the StatProgressBar example

'
' API Declarations
'
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any) As Long

'
' API Types
'
' RECT is used to get the size of the panel we're inserting into
'
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'
' API Messages
'
Public Const WM_USER As Long = &H400
Public Const SB_GETRECT As Long = (WM_USER + 10)
