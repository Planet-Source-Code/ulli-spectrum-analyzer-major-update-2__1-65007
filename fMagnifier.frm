VERSION 5.00
Begin VB.Form fMagnifier 
   Appearance      =   0  '2D
   AutoRedraw      =   -1  'True
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   690
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   660
   ControlBox      =   0   'False
   DrawMode        =   14  'Stift kopieren
   ForeColor       =   &H00000000&
   Icon            =   "fMagnifier.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   46
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   44
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer tmrRepaint 
      Interval        =   30
      Left            =   195
      Top             =   180
   End
End
Attribute VB_Name = "fMagnifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'###############################################################
#Const Rounded = True                    'set to False for square viewport
Private Const Magnify       As Long = 5  'magnifying factor
Private Const ViewSize      As Long = 40 'size of area to be magnified (in pixels)
'###############################################################

Private Const ViewSize2     As Long = ViewSize \ 2
Private Const DestSize      As Long = ViewSize * Magnify
Private Const DestCenter    As Long = DestSize \ 2
Private Const CrossCenter   As Long = DestCenter + Magnify \ 2 + (Magnify And 1) 'crosshair at center of magnified pixel

Private Const Border        As Long = 1
Private ScrXMax             As Long
Private ScrYMax             As Long

Private Type tPoint
    X   As Long
    Y   As Long
End Type
Private CursorPos           As tPoint
Private xPos                As Long
Private yPos                As Long

'handles and api constans
Private Const hWndDesktop   As Long = 0
Private hDCDesktop          As Long
Private hRgn                As Long
Private Const HWND_TOPMOST  As Long = -1
Private Const SWP_NOSIZE    As Long = 1
Private Const SWP_NOMOVE    As Long = 2

Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As tPoint) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDstDC As Long, ByVal xDst As Long, ByVal yDst As Long, ByVal nDstWidth As Long, ByVal nDtsHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Sub Form_Load()

    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    hDCDesktop = GetWindowDC(hWndDesktop)
    Width = ScaleX(DestSize + Border + Border, vbPixels, vbTwips)
    Height = Width
    Show
    Repaint

#If Rounded Then '------------------------------------------------------------------
    xPos = ScaleWidth
    hRgn = CreateRoundRectRgn(2, 2, xPos, xPos, xPos, xPos)
    SetWindowRgn hWnd, hRgn, True
#End If '---------------------------------------------------------------------------

    With Screen
        ScrXMax = ScaleX(.Width, vbTwips, vbPixels) - ViewSize
        ScrYMax = ScaleY(.Height, vbTwips, vbPixels) - ViewSize
    End With 'SCREEN

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If hDCDesktop Then
        ReleaseDC hWndDesktop, hDCDesktop
    End If
    If hRgn Then
        DeleteObject hRgn
    End If

End Sub

Public Sub Repaint()

    GetCursorPos CursorPos
    With CursorPos
        xPos = ScaleX(.X + ViewSize2, vbPixels, vbTwips)
        yPos = ScaleY(.Y + ViewSize2, vbPixels, vbTwips)
        If xPos + Width > Screen.Width Then
            xPos = Screen.Width - Width
        End If
        If yPos + Height > Screen.Height Then
            yPos = Screen.Height - Height
        End If
        If xPos = Screen.Width - Width And yPos = Screen.Height - Height Then
            xPos = ScaleX(.X - ViewSize2, vbPixels, vbTwips) - Width
            yPos = ScaleY(.Y - ViewSize2, vbPixels, vbTwips) - Height
        End If
        Move xPos, yPos
        xPos = .X - ViewSize2
        yPos = .Y - ViewSize2
        Cls
        StretchBlt hDC, Border, Border, DestSize, DestSize, hDCDesktop, xPos, yPos, ViewSize, ViewSize, vbSrcCopy
        Line (CrossCenter, CrossCenter - 7)-(CrossCenter, CrossCenter + 8), vbRed
        Line (CrossCenter - 7, CrossCenter)-(CrossCenter + 8, CrossCenter), vbRed

#If Rounded Then '------------------------------------------------------------------
        DrawWidth = 2
        DrawMode = vbCopyPen
        Circle (DestCenter - 1, DestCenter - 1), DestCenter - 2, &H808000
        DrawWidth = 1
        DrawMode = vbMergePenNot
#End If '---------------------------------------------------------------------------
    End With 'CURSORPOS

End Sub

':) Ulli's VB Code Formatter V2.21.6 (2006-Apr-24 11:50)  Decl: 41  Code: 71  Total: 112 Lines
':) CommentOnly: 3 (2,7%)  Commented: 10 (8,9%)  Empty: 18 (16,1%)  Max Logic Depth: 3
