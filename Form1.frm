VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   2655
   StartUpPosition =   2  'CenterScreen
   Begin DrawColoredTextToScreen.TransMsg TransMsg1 
      Height          =   360
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   360
      _extentx        =   635
      _extenty        =   635
   End
   Begin VB.CheckBox chkBold 
      Caption         =   "Bold"
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   1320
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Text            =   "14"
      Top             =   1320
      Width           =   375
   End
   Begin MSComDlg.CommonDialog cmDlg 
      Left            =   1080
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdColor 
      BackColor       =   &H00EFEFEF&
      Caption         =   "..."
      Height          =   600
      Left            =   1845
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "set text color"
      Top             =   585
      Width           =   600
   End
   Begin VB.OptionButton optAlign 
      Caption         =   "Right align"
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   6
      Top             =   945
      Width           =   1680
   End
   Begin VB.OptionButton optAlign 
      Caption         =   "Center"
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   5
      Top             =   765
      Value           =   -1  'True
      Width           =   1680
   End
   Begin VB.OptionButton optAlign 
      Caption         =   "Left align"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   4
      Top             =   585
      Width           =   1680
   End
   Begin VB.TextBox txtDisplay 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H002F2F2F&
      Height          =   2000
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   2040
      Width           =   2500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Erase text"
      Height          =   420
      Index           =   1
      Left            =   1260
      TabIndex        =   1
      Top             =   45
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Draw text"
      Height          =   420
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   45
      Width           =   1095
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "text message to display"
      ForeColor       =   &H00A00000&
      Height          =   405
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   2505
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

       txtDisplay = _
 _
"This is a really cool way to display a message" & vbCrLf & _
"to the user without the intrusiveness that" & vbCrLf & _
"a message box can cause." & vbCrLf & vbCrLf & _
"Its like a mini help file you can display and then" & vbCrLf & _
"erase at will." & vbCrLf & _
"You can display it anywhere, and anytime " & vbCrLf & _
"you wish!!"
TransMsg1.TXT_COLOR = vbRed
End Sub

Private Sub cmdColor_Click()
On Error GoTo ERR:
   cmDlg.CancelError = True
   cmDlg.ShowColor
   TransMsg1.TXT_COLOR = cmDlg.Color
Exit Sub
ERR:
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Dim lng&
    Case Is = 0 'draw the text
        '   If optAlign(0).Value = True Then
        '       lng = DT_left
        '   ElseIf optAlign(1).Value = True Then
        '       lng = DT_CENTER
        '   ElseIf optAlign(2).Value = True Then
        '       lng = DT_right
        '   End If
        'this is the bounding rectangle
        'if the text gets cliped, just increase
        'its width or height
        TransMsg1.DrawTextToScreen txtDisplay.Text, _
                                       Screen.Width * 0.6, _
                                       Screen.Width * 0.8, _
                                       Screen.Height * 0.3, _
                                       Screen.Height * 0.6, _
                                       lng, _
                                       TransMsg1.TXT_COLOR, _
                                       CBool(chkBold.Value), _
                                       txt1

    Case Is = 1 'erase the text
            TransMsg1.EraseTextDrawnToScreen
End Select
End Sub
