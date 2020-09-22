VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmStatProgBar 
   BackColor       =   &H00A07247&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Status / Progress Bar Example"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3360
      Top             =   30
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   300
      Left            =   15
      TabIndex        =   2
      Top             =   585
      Visible         =   0   'False
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   529
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Progress Bar"
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   45
      Width           =   1545
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   1245
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Text            =   "Example Panel"
            TextSave        =   "Example Panel"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   4286
            Text            =   "Progress Bar Goes Here....."
            TextSave        =   "Progress Bar Goes Here....."
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "10:02 AM"
            Key             =   "ProgBar"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmStatProgBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This form demonstrates how to place a Progress-Bar into a
' panel of a status bar.
'

Private Sub Command1_Click()
'
' Disable this button for now
'
    Command1.Enabled = False
'
' Setup the progress bar with some values
'
    ProgressBar1.Min = 0
    ProgressBar1.Max = 100
'
' Show ProgressBar in Status Bar
'
    ShowProgressInStatusBar True
'
' Enable the timer so it looks like we're doing something
'
    Timer1.Enabled = True
End Sub

Private Sub ShowProgressInStatusBar(ByVal bShowProgressBar As Boolean)

    Dim tRC As RECT
    
    If bShowProgressBar Then
'
' Get the size of the Panel (2) Rectangle from the status bar
' remember that Indexes in the API are always 0 based (well,
' nearly always) - therefore Panel(2) = Panel(1) to the api
'
'
        SendMessageAny StatusBar1.hwnd, SB_GETRECT, 1, tRC
'
' and convert it to twips....
'
        With tRC
            .Top = (.Top * Screen.TwipsPerPixelY)
            .Left = (.Left * Screen.TwipsPerPixelX)
            .Bottom = (.Bottom * Screen.TwipsPerPixelY) - .Top
            .Right = (.Right * Screen.TwipsPerPixelX) - .Left
        End With
'
' Now Reparent the ProgressBar to the statusbar
'
        With ProgressBar1
            SetParent .hwnd, StatusBar1.hwnd
            .Move tRC.Left, tRC.Top, tRC.Right, tRC.Bottom
            .Visible = True
            .Value = 0
        End With
        
    Else
'
' Reparent the progress bar back to the form and hide it
'
        SetParent ProgressBar1.hwnd, Me.hwnd
        ProgressBar1.Visible = False
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
' Should really re-parent the progress bar here,
' just in case anything terrible happened
'
    ShowProgressInStatusBar False
    
End Sub

Private Sub Timer1_Timer()
'
' This timer routine simply updates the progress bar to make it
' seem like there's something going on....
'
    Static lCount As Long
    
    lCount = lCount + 5
    
    If lCount > 100 Then
        Timer1.Enabled = False
        ShowProgressInStatusBar False
        Command1.Enabled = True
        lCount = 0
    End If
    
    ProgressBar1.Value = lCount
    
End Sub
