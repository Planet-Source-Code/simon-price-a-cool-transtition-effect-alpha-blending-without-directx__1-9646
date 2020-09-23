VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cool Transition by Simon Price"
   ClientHeight    =   5868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   5808
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   489
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   484
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog ComDialog 
      Left            =   2520
      Top             =   120
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
      Filter          =   "*.bmp, *.jpg, *.jpeg"
   End
   Begin VB.PictureBox Display 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   2412
      Left            =   3000
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   11
      Top             =   480
      Width           =   2412
   End
   Begin VB.CommandButton cmdGO 
      Caption         =   "GO !!!"
      Height          =   492
      Left            =   3720
      TabIndex        =   8
      Top             =   5280
      Width           =   1212
   End
   Begin VB.HScrollBar SpeedScroll 
      Height          =   372
      Left            =   2760
      Max             =   10
      Min             =   1
      TabIndex        =   6
      Top             =   4680
      Value           =   7
      Width           =   2892
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse ..."
      Height          =   372
      Index           =   2
      Left            =   1440
      TabIndex        =   3
      Top             =   2880
      Width           =   972
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse ..."
      Height          =   372
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   972
   End
   Begin VB.PictureBox TestPic 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   2412
      Index           =   2
      Left            =   120
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   9
      Top             =   3240
      Width           =   2412
   End
   Begin VB.PictureBox TestPic 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   2412
      Index           =   1
      Left            =   120
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   10
      Top             =   480
      Width           =   2412
   End
   Begin VB.PictureBox PB 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   2412
      Left            =   5040
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   2412
   End
   Begin VB.Label TimeLabel 
      Caption         =   "Speed of Transition : (Slow ---- Fast)"
      Height          =   372
      Left            =   2760
      TabIndex        =   7
      Top             =   4320
      Width           =   3012
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Form1.frx":030A
      Height          =   1200
      Left            =   2760
      TabIndex        =   5
      Top             =   3000
      Width           =   2868
   End
   Begin VB.Label Label2 
      Caption         =   "Cool Transition Here :"
      Height          =   252
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   1932
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Image 2 :"
      Height          =   192
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   648
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Image 1 :"
      Height          =   192
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   648
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'COOL TRANSITION BY SIMON PRICE

'declarations
'get pixel for looking at pixel values
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
'set pixel for drawing pixels
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
'bitblt for copy from my back buffer
'(a picturebox) to the front buffer
'(another picturebox)
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'colours can be long or red-green-blue, so I made
'a type called RGBcolor to store each of the 3
'byte of info
Private Type RGBcolor
  R As Byte 'amount of red
  G As Byte 'amount of green
  B As Byte 'amount of blue
End Type

'this is how many steps the transition makes to
'get from image 1 to image 2
Public Steps As Byte

'stores the file paths of the chosen pictures
Dim PicPath(1 To 2) As String

Private Sub cmdBrowse_Click(Index As Integer)
'show open picture dialog box
ComDialog.ShowOpen
'set file name
PicPath(Index) = ComDialog.FileName
'load the pictures again
LoadPictures
End Sub

Private Sub cmdGO_Click()
'do the transition when the GO button is clicked
DoTransition
End Sub

Private Sub Form_Load()
'load the defualt pictures
LoadPictures
End Sub

Sub LoadPictures()
'loads the pictures into the picture boxes
On Error GoTo YouMuckedItUp
'if there is no file path, the default pictures are loaded instead
If PicPath(1) = "" Then PicPath(1) = App.Path & "\some_geezer.bmp"
If PicPath(2) = "" Then PicPath(2) = App.Path & "\my_jump.bmp"
'load pictures
TestPic(1) = LoadPicture(PicPath(1))
TestPic(2) = LoadPicture(PicPath(2))
Display = TestPic(1)
Exit Sub
'if there's an error then display a message box
YouMuckedItUp:
MsgBox "There was an error when attempting to load the pictures", vbCritical, "Error!"
End Sub

Sub DoTransition()
'this is the cool bit, the transition is done here
On Error Resume Next
'used for the 'for' loops
Dim x, y, i As Byte
'these are step values, each colour increases by
'the step value each time it changes
Dim StepR(0 To 200, 0 To 200) As Single
Dim StepG(0 To 200, 0 To 200) As Single
Dim StepB(0 To 200, 0 To 200) As Single
'these store the RGB of every pixel
Dim R(0 To 200, 0 To 200) As Single
Dim G(0 To 200, 0 To 200) As Single
Dim B(0 To 200, 0 To 200) As Single
'temporary long value for loading colours
Dim TempLong As Long
'the initial RGB colour of a pixel
Dim StartCol As RGBcolor
'the final RGB colour of a pixel
Dim EndCol As RGBcolor
'these store the difference between the start and end colours
Dim DiffColR As Integer
Dim DiffColG As Integer
Dim DiffColB As Integer

'in the loading stage, change the mousepointer
Me.MousePointer = vbHourglass

'the number of steps for the transition
'more steps = quality transition = slow
'less steps = cheap transition = fast
Steps = (11 - SpeedScroll.Value) * 10

'now loop through every pixel and find the difference between
'the start and end values, and then the step values
For x = 0 To 200
For y = 0 To 200
  
  'get the RGB value of the start pixel
  TempLong = GetPixel(TestPic(1).hdc, x, y)
  StartCol.R = TempLong And 255
  StartCol.G = (TempLong And 65280) \ 256&
  StartCol.B = (TempLong And 16711680) \ 65535
  
  'get the RGB value of the end pixel
  TempLong = GetPixel(TestPic(2).hdc, x, y)
  EndCol.R = TempLong And 255
  EndCol.G = (TempLong And 65280) \ 256&
  EndCol.B = (TempLong And 16711680) \ 65535
  
  'set initial RGB values
  R(x, y) = StartCol.R
  G(x, y) = StartCol.G
  B(x, y) = StartCol.B
  
  'work out the difference between the start and end red values
  If EndCol.R > StartCol.R Then
    DiffColR = EndCol.R - StartCol.R
  Else
    DiffColR = StartCol.R - EndCol.R
    DiffColR = -DiffColR
  End If
  
  'work out the difference between the start and end green values
  If EndCol.G > StartCol.G Then
    DiffColG = EndCol.G - StartCol.G
  Else
    DiffColG = StartCol.G - EndCol.G
    DiffColG = -DiffColG
  End If
  
  'work out the difference between the start and end blue values
  If EndCol.B > StartCol.B Then
    DiffColB = EndCol.B - StartCol.B
  Else
    DiffColB = StartCol.B - EndCol.B
    DiffColB = -DiffColB
  End If
  
  'work out the step value by dividing the difference by the
  'number of steps
  StepR(x, y) = DiffColR / Steps
  StepG(x, y) = DiffColG / Steps
  StepB(x, y) = DiffColB / Steps
  
Next
Next

'loading is finished, so change the mousepointer
'back to normal
Me.MousePointer = vbDefault

'now loop through and draw each step of the transition
For i = 1 To Steps
    'loop through every pixel
    For x = 0 To 200
    For y = 0 To 200
        'increase the RGB values by the step values
        'we worked out earlier
        R(x, y) = R(x, y) + StepR(x, y)
        G(x, y) = G(x, y) + StepG(x, y)
        B(x, y) = B(x, y) + StepB(x, y)
        'draw the pixel on the invisible picturebox
        SetPixel PB.hdc, x, y, RGB(R(x, y), G(x, y), B(x, y))
    Next
    Next
'the step is done, so copy from the invisible
'picturebox the the visible one on display
BitBlt Display.hdc, 0, 0, 200, 200, PB.hdc, 0, 0, vbSrcCopy
Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "If you thought that that was a cool effect, then please vote for me on Planet Source Code!!!", , "Vote Now! - Cool Transition by Simon Price"
End Sub
