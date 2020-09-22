VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form KaleidescopeFrm 
   Caption         =   "Kaleidescope"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   332
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton LoadButton 
      Caption         =   "Load Pic"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog DLG 
      Left            =   120
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.HScrollBar RotScroll 
      Height          =   255
      LargeChange     =   2
      Left            =   120
      Max             =   10
      Min             =   -10
      TabIndex        =   4
      Top             =   2280
      Value           =   5
      Width           =   1695
   End
   Begin VB.PictureBox ViewImage 
      Height          =   3000
      Left            =   1920
      Picture         =   "KaleidescopeFrm.frx":0000
      ScaleHeight     =   196
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   196
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Timer ROF 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   2760
   End
   Begin VB.CommandButton StopButton 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton StartButton 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Degrees/Frame:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Frame Rate:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   870
   End
   Begin VB.Label RotLbl 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label ROFLbl 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   1560
      Width           =   375
   End
   Begin VB.Image TextureImage 
      Height          =   3000
      Left            =   1920
      Picture         =   "KaleidescopeFrm.frx":A082
      Top             =   120
      Width           =   3000
   End
End
Attribute VB_Name = "KaleidescopeFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const PI2 = 6.28318530717959
Dim done As Boolean
Dim pieX() As Integer 'coordinates of a pie slice
Dim pieY() As Integer
Dim nPie As Long 'number of elements in above tables
Dim cosarray(359) As Single
Dim sinarray(359) As Single
Dim Rot As Integer 'rotation rate
Dim BMPSIZE As Integer  'width or height of a SQUARE bitmap, in pixels
Dim HALFBMPSIZE As Integer
Dim Rate As Integer

Private Sub LoadPies()
'loads the pie arrays
Dim x%, y%
Dim x1#, y1#
BMPSIZE = TextureImage.Width
HALFBMPSIZE = BMPSIZE \ 2
nPie = 0
ReDim pieX(7, 0) As Integer
ReDim pieY(7, 0) As Integer
For x% = HALFBMPSIZE To BMPSIZE - 1
    For y% = HALFBMPSIZE To BMPSIZE - 1
        x1 = x% - HALFBMPSIZE + 1
        y1 = y - HALFBMPSIZE + 1
        If (x1 * x1 + y1 * y1) < (HALFBMPSIZE * CLng(HALFBMPSIZE)) Then
            If Abs(Atn(y1 / x1)) < (PI2 / 8) Then
                ReDim Preserve pieX(7, nPie) As Integer
                ReDim Preserve pieY(7, nPie) As Integer
                pieX(0, nPie) = x
                pieY(0, nPie) = y
                pieX(1, nPie) = x
                pieY(1, nPie) = BMPSIZE - y - 1
                pieX(2, nPie) = BMPSIZE - x - 1
                pieY(2, nPie) = y
                pieX(3, nPie) = BMPSIZE - x - 1
                pieY(3, nPie) = BMPSIZE - y - 1
                pieX(4, nPie) = y
                pieY(4, nPie) = x
                pieX(5, nPie) = y
                pieY(5, nPie) = BMPSIZE - x - 1
                pieX(6, nPie) = BMPSIZE - y - 1
                pieY(6, nPie) = x
                pieX(7, nPie) = BMPSIZE - y - 1
                pieY(7, nPie) = BMPSIZE - x - 1
                nPie = nPie + 1
            End If
        End If
    Next y%
Next x%
nPie = nPie - 1
End Sub

Private Sub Form_Load()
'mostly load arrays

Dim x%

For x% = 0 To 359
    cosarray(x%) = Cos(CDbl(x%) * PI2 / 360)
    sinarray(x%) = Sin(CDbl(x%) * PI2 / 360)
Next x%
LoadPies
Rot = RotScroll.Value

LoadPicArray2D TextureImage.Picture, TextureSA, TextureBMP, TextureData()
LoadPicArray2D ViewImage.Picture, ViewSA, ViewBMP, ViewData()

End Sub

Private Sub Form_Unload(Cancel As Integer)
done = True
PicArrayKill TextureData()
PicArrayKill ViewData()
End
End Sub

Private Sub mainloop()
Static Theta As Integer 'main angle
Dim x As Single, y As Single
Dim i&
Dim j%, x1%
Dim c As Byte

Do
'adjust rotation
    Theta = (Theta + Rot + 360) Mod 360
'cut out section of texture map and paste
    For i& = 0 To nPie
        x1% = pieX(0, i&) - HALFBMPSIZE
        y = pieY(0, i&) - HALFBMPSIZE
        x = x1% * cosarray(Theta) - y * sinarray(Theta) + HALFBMPSIZE
        y = y * cosarray(Theta) + x1% * sinarray(Theta) + HALFBMPSIZE
        c = TextureData(x, y)
        For j% = 0 To 7
            ViewData(pieX(j%, i&), pieY(j%, i&)) = c
        Next
    Next
    ViewImage.Refresh
    Rate = Rate + 1
    DoEvents
Loop While Not done
End Sub

Private Sub LoadButton_Click()
On Error Resume Next
Dim s$
Dim i%, j%

DLG.CancelError = True
DLG.InitDir = App.Path
DLG.DefaultExt = "BMP"
DLG.FileName = "*.bmp"
DLG.Filter = "Bitmaps|*.bmp"
DLG.ShowOpen
If Err = 0 Then
    s$ = DLG.FileName
    TextureImage.Picture = LoadPicture(s$)
    ViewImage.Picture = LoadPicture(s$)
    TextureImage.Refresh
    ViewImage.Width = TextureImage.Width
    ViewImage.Height = TextureImage.Height
    i% = Me.Width - Me.ScaleWidth * Screen.TwipsPerPixelX
    j% = Me.Height - Me.ScaleHeight * Screen.TwipsPerPixelY
    Me.Width = (ViewImage.Left + ViewImage.Width + StartButton.Left) * Screen.TwipsPerPixelX + i%
    Me.Height = (ViewImage.Top + ViewImage.Height + StartButton.Top) * Screen.TwipsPerPixelY + j%
    LoadPicArray2D TextureImage.Picture, TextureSA, TextureBMP, TextureData()
    LoadPicArray2D ViewImage.Picture, ViewSA, ViewBMP, ViewData()
    LoadPies
    For i% = 0 To BMPSIZE - 1
        For j% = 0 To BMPSIZE - 1
            ViewData(i%, j%) = 0
        Next j%
    Next i%
    ViewImage.Refresh
End If
End Sub

Private Sub ROF_Timer()
ROFLbl = CStr(Rate)
Rate = 0
End Sub

Private Sub RotScroll_Change()
Rot = RotScroll.Value
RotLbl = CStr(Rot)
End Sub

Private Sub StartButton_Click()
done = False
ROF.Enabled = True
ViewImage.Visible = True
TextureImage.Visible = False
StopButton.Enabled = True
StartButton.Enabled = False
mainloop
End Sub

Private Sub StopButton_Click()
done = True
ROF.Enabled = False
ViewImage.Visible = False
TextureImage.Visible = True
StartButton.Enabled = True
StopButton.Enabled = False
End Sub

