VERSION 5.00
Begin VB.UserControl TransMsg 
   ClientHeight    =   8610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10605
   ScaleHeight     =   8610
   ScaleWidth      =   10605
   ToolboxBitmap   =   "TransMsg.ctx":0000
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   7335
      Left            =   480
      ScaleHeight     =   7275
      ScaleWidth      =   10035
      TabIndex        =   0
      Top             =   0
      Width           =   10095
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   0
      Picture         =   "TransMsg.ctx":0312
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "TransMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const DT_WORDBREAK = &H10

Enum TextAlign
    DT_left = &H0
    DT_right = &H2
    DT_CENTER = &H1
End Enum

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private dec_deskMsgRect As RECT
Public TXT_COLOR As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
                                             ByVal x As Long, _
                                             ByVal Y As Long, _
                                             ByVal nWidth As Long, _
                                             ByVal nHeight As Long, _
                                             ByVal hSrcDC As Long, _
                                             ByVal xSrc As Long, _
                                             ByVal ySrc As Long, _
                                             ByVal dwRop As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
                                                ByVal nWidth As Long, _
                                                ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function InvalidateRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
 
Public Sub DrawTextToScreen(sDeskMsg$, iTextLeft%, iTextRight%, iTextTop%, iTextBottom%, _
                                  Optional TXT_ALIGN As TextAlign = DT_CENTER, _
                                  Optional TxtColor& = vbBlack, _
                                  Optional FBold As Boolean = True, _
                                  Optional FSize As Integer)

Dim screenHdc As Long

'erase text before drawing new
Call EraseTextDrawnToScreen
DoEvents

' Get hDC of Desktop
screenHdc = GetDC(0)

SetBkMode Picture1.hdc, 3&

BitBlt Picture1.hdc, -200, -200, Picture1.ScaleWidth, Picture1.ScaleHeight, screenHdc, 0, 0, vbSrcCopy

Picture1.Refresh

'set the textcolor
SetTextColor Picture1.hdc, TxtColor

Picture1.FontBold = FBold
Picture1.FontSize = FSize

'rectangle where we draw the text
SetRect dec_deskMsgRect, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight

DrawText Picture1.hdc, sDeskMsg, Len(sDeskMsg), _
                              dec_deskMsgRect, _
                              DT_WORDBREAK Or DT_left ' TXT_ALIGN
                              
Picture1.Refresh
'picture1 to screen
BitBlt screenHdc, 200, 200, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hdc, 0, 0, vbSrcCopy
End Sub

Public Sub EraseTextDrawnToScreen()
InvalidateRect 0, dec_deskMsgRect, True
End Sub

Private Sub UserControl_Resize()
UserControl.Width = 360
UserControl.Height = 360
End Sub
