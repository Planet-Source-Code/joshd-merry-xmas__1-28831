VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   825
   ClientTop       =   915
   ClientWidth     =   12000
   DrawWidth       =   5
   BeginProperty Font 
      Name            =   "Monotype Corsiva"
      Size            =   80.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picScreen 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9000
      Left            =   2760
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   1
      Top             =   4320
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.PictureBox picMap 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9000
      Left            =   1560
      Picture         =   "Form1.frx":36025
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   -480
      Visible         =   0   'False
      Width           =   12000
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NewX As Integer
Dim NewY As Integer
Dim I As Integer
Dim J As Integer
Dim K As Integer
Dim XMas As Integer
Dim XMasChange As Integer
Dim SCREEN_WIDTH As Integer
Dim SCREEN_HEIGHT As Integer
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Sub Form_Activate()
Me.Show
SCREEN_WIDTH = 800
SCREEN_HEIGHT = 600
Do
    DoSnow
    DrawStuff
    DoEvents
Loop
End Sub
Public Sub DoSnow()
    Randomize
    picMap.DrawWidth = 1
    picScreen.DrawWidth = 1
    For I = 1 To 1000
        If Snow(I).XPos > SCREEN_WIDTH Then
            Snow(I).XPos = 0
        End If
        If Snow(I).XPos < 0 Then
            Snow(I).XPos = SCREEN_WIDTH
        End If
        If Snow(I).YPos > SCREEN_HEIGHT Then
            Snow(I).YPos = 0
        End If
        If Snow(I).Used = True Then
            If GetPixel(picMap.hdc, Snow(I).XPos + Snow(I).XSpeed, Snow(I).YPos + Snow(I).YSpeed) <> 0 Then
                Snow(I).YPos = Snow(I).YPos + Snow(I).YSpeed
                Snow(I).XPos = Snow(I).XPos + Snow(I).XSpeed

            ElseIf Snow(I).YSpeed = 1 Then  'The snow is slow, ie going down
                Dim vSolved As Boolean
                vSolved = False
                For J = 6 To 1 Step -1
                    If GetPixel(picMap.hdc, Snow(I).XPos + J, Snow(I).YPos + 1) <> 0 And (Rnd * 7) > J Then
                        Snow(I).YPos = Snow(I).YPos + 1
                        Snow(I).XPos = Snow(I).XPos + J
                        vSolved = True
                        Exit For
                    End If
                Next J
                If Not vSolved Then
                    Snow(I).Used = False
                    Call SetPixel(picMap.hdc, Snow(I).XPos, Snow(I).YPos, vbBlack)
                    Call SetPixel(picScreen.hdc, Snow(I).XPos, Snow(I).YPos, snowColour)
                End If
            Else
                Snow(I).YSpeed = 1
                Snow(I).XSpeed = 0
            End If
        ElseIf Rnd * 1000 < 2 Then
            Snow(I).Used = True
            Snow(I).XPos = Rnd * 798 + 1
            Snow(I).YPos = 0
            Snow(I).YSpeed = 1.5 + (Rnd * 1)
            Snow(I).XSpeed = (1 + Snow(I).YSpeed) / 4
        End If
    Next I
End Sub
Public Function snowColour()
    Randomize
    Dim rndNum As Integer
    rndNum = 255 '- Rnd * 50
    snowColour = RGB(rndNum, rndNum, rndNum)
End Function
Public Sub DrawStuff()
    Dim colour As Long
    Dim col As Long
    Call BitBlt(frmMain.hdc, 0, 0, 800, 600, picScreen.hdc, 0, 0, SRCCOPY)
    Call BitBlt(frmMain.hdc, 0, 0, 800, 600, picScreen.hdc, 0, 0, SRCCOPY)
    Call BitBlt(frmMain.hdc, 0, 0, 800, 600, picScreen.hdc, 0, 0, SRCCOPY)
    picMap.DrawWidth = 4
    picScreen.DrawWidth = 4
    For I = 1 To 1000
        If Snow(I).Used = True Then
            If GetPixel(picMap.hdc, Snow(I).XPos, Snow(I).YPos) <> 0 And Snow(I).XPos < picMap.Width And Snow(I).XPos > 0 Then
                col = Snow(I).YSpeed * 100
                colour = RGB(col, col, col)
                Call SetPixel(frmMain.hdc, Snow(I).XPos, Snow(I).YPos, colour)
            End If
        End If
    Next I
    picMap.DrawWidth = 1
    picScreen.DrawWidth = 1
    
    
    XMas = XMas + XMasChange
    If XMas > 255 Then
        XMasChange = XMasChange * -1
        XMas = 255
    ElseIf XMas < 0 Then
        XMasChange = XMasChange * -1
        XMas = 0
    End If
    frmMain.ForeColor = RGB(XMas, 255 - XMas, 0)
    frmMain.CurrentX = (SCREEN_WIDTH - frmMain.TextWidth("Merry Christmas")) / 2
    frmMain.CurrentY = -10
    frmMain.Print "Merry Christmas"
    
    frmMain.Refresh
End Sub

Private Sub Form_Load()
    XMas = 0
    XMasChange = 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

