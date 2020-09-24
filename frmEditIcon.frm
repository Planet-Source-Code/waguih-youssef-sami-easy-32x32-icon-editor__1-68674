VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditIcon 
   Caption         =   "Edit Icon"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   Icon            =   "frmEditIcon.frx":0000
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   505
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00AB8F8D&
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   6360
      ScaleHeight     =   4785
      ScaleWidth      =   1185
      TabIndex        =   24
      Top             =   1200
      Width           =   1215
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         DrawWidth       =   2
         ForeColor       =   &H80000008&
         Height          =   3825
         Left            =   360
         ScaleHeight     =   3795
         ScaleWidth      =   435
         TabIndex        =   25
         Top             =   120
         Width           =   465
      End
      Begin VB.Label Pallet2 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   4080
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   120
         Top             =   3960
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   7545
      TabIndex        =   20
      Top             =   120
      Width           =   7575
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3945
         Top             =   15
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   3315
         Top             =   30
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   15
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditIcon.frx":030A
               Key             =   "Open"
               Object.Tag             =   "Open"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditIcon.frx":09DE
               Key             =   "Save"
               Object.Tag             =   "Save"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditIcon.frx":0E30
               Key             =   "New"
               Object.Tag             =   "New"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditIcon.frx":0F8A
               Key             =   "Fill"
               Object.Tag             =   "Fill"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditIcon.frx":10E4
               Key             =   "Paint"
               Object.Tag             =   "Paint"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditIcon.frx":123E
               Key             =   "Down"
               Object.Tag             =   "Down"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditIcon.frx":1398
               Key             =   "Up"
               Object.Tag             =   "Up"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditIcon.frx":14F2
               Key             =   "Right"
               Object.Tag             =   "Right"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditIcon.frx":164C
               Key             =   "Left"
               Object.Tag             =   "Left"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditIcon.frx":17A6
               Key             =   "Refresh"
               Object.Tag             =   "Refresh"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditIcon.frx":1900
               Key             =   "Exit"
               Object.Tag             =   "Exit"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditIcon.frx":1A5A
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditIcon.frx":1BB4
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditIcon.frx":1D0E
               Key             =   "Eraser"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditIcon.frx":202A
               Key             =   "Grab"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   2640
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   12648447
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   16777215
         _Version        =   393216
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Waguih Icon Editor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   3495
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00AB8F8D&
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   120
      ScaleHeight     =   4785
      ScaleWidth      =   1185
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
      Begin VB.PictureBox SmallPicture 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   360
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   2
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   480
         Left            =   360
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   480
      End
      Begin VB.Label lblColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   15
         Left            =   240
         TabIndex        =   19
         Top             =   2760
         Width           =   225
      End
      Begin VB.Label lblColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   14
         Left            =   720
         TabIndex        =   18
         Top             =   2520
         Width           =   225
      End
      Begin VB.Label Pallet 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   480
         TabIndex        =   17
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label lblColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   13
         Left            =   480
         TabIndex        =   16
         Top             =   2520
         Width           =   225
      End
      Begin VB.Label lblColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   12
         Left            =   240
         TabIndex        =   15
         Top             =   2520
         Width           =   225
      End
      Begin VB.Label lblColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   11
         Left            =   720
         TabIndex        =   14
         Top             =   2280
         Width           =   225
      End
      Begin VB.Label lblColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   10
         Left            =   480
         TabIndex        =   13
         Top             =   2280
         Width           =   225
      End
      Begin VB.Label lblColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   9
         Left            =   240
         TabIndex        =   12
         Top             =   2280
         Width           =   225
      End
      Begin VB.Label lblColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   8
         Left            =   720
         TabIndex        =   11
         Top             =   2040
         Width           =   225
      End
      Begin VB.Label lblColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   7
         Left            =   480
         TabIndex        =   10
         Top             =   2040
         Width           =   225
      End
      Begin VB.Label lblColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   6
         Left            =   240
         TabIndex        =   9
         Top             =   2040
         Width           =   225
      End
      Begin VB.Label lblColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   5
         Left            =   720
         TabIndex        =   8
         Top             =   1800
         Width           =   225
      End
      Begin VB.Label lblColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   4
         Left            =   480
         TabIndex        =   7
         Top             =   1800
         Width           =   225
      End
      Begin VB.Label lblColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   225
      End
      Begin VB.Label lblColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   720
         TabIndex        =   5
         Top             =   1560
         Width           =   225
      End
      Begin VB.Label lblColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   480
         TabIndex        =   4
         Top             =   1560
         Width           =   225
      End
      Begin VB.Label lblColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   225
      End
   End
   Begin VB.PictureBox LargePicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   1440
      ScaleHeight     =   319
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   319
      TabIndex        =   0
      Top             =   1200
      Width           =   4815
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   0
      TabIndex        =   22
      Top             =   840
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Open Icon"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Open  Bitmap"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageIndex      =   3
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   2
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Save_Icon"
                  Text            =   "Save As Icon"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Save_BMP"
                  Text            =   "Save As BMP"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paint"
            Object.ToolTipText     =   "Paint"
            ImageIndex      =   5
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Change"
            Object.ToolTipText     =   "Change"
            ImageIndex      =   4
            Style           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Left"
            Object.ToolTipText     =   "Move Left"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Down"
            Object.ToolTipText     =   "Move Down"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Up"
            Object.ToolTipText     =   "Move Up"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Right"
            Object.ToolTipText     =   "Move Right"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "flipHorizontal"
            Object.ToolTipText     =   "Flip Horizontal"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FlipVert"
            Object.ToolTipText     =   "Flip Vertically"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Rotate"
            Object.ToolTipText     =   "Rotate"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Clear"
            Object.ToolTipText     =   "Clear All"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Extract"
            Object.ToolTipText     =   "Extract Icon"
            ImageIndex      =   15
            Value           =   1
         EndProperty
      EndProperty
   End
   Begin VB.Label lblStatusBar 
      BackColor       =   &H00AB8F8D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   6120
      Width           =   7455
   End
End
Attribute VB_Name = "frmEditIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyColor As Long, b As Long, c As Long, d As Long, e As Long
Dim GridFormat As Integer, ValAdd As Integer, upPoint As Integer
Dim Picture0 As StdPicture

Private Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Const DI_MASK = &H1
Private Const DI_IMAGE = &H2
Private Const DI_NORMAL = DI_MASK Or DI_IMAGE
Private TotalApps As Integer


Private Sub Form_Load()

'***************Initiate some Variables*************
GridFormat = 32
upPoint = 9
MyColor = vbBlack
Dim GridStep As Integer
GridStep = 10
ValAdd = 10
'******************Fill the Pallet with Colors*********
'    For e = 0 To 15
'      lblColor(e).BackColor = QBColor(e)
'    Next e

lblColor(0).BackColor = vbWindowBackground '&H80000005
lblColor(1).BackColor = vbGreen             '&HFF00
lblColor(2).BackColor = vbDesktop           '&H80000001
lblColor(3).BackColor = vbCyan              '&HFFFF00
lblColor(4).BackColor = vbBlue              '&HFF0000
lblColor(5).BackColor = vbInfoBackground    '&H80000018
lblColor(6).BackColor = vbYellow            '&HFFFF
lblColor(7).BackColor = vb3DHighlight       '&H80000014
lblColor(8).BackColor = vbInactiveBorder    '&H8000000B
lblColor(9).BackColor = vbInactiveTitleBar  '&H80000003
lblColor(10).BackColor = vbHighlight        '&H8000000D
lblColor(11).BackColor = vb3DLight          '&H80000016
lblColor(12).BackColor = vbBlack            '&H0
lblColor(13).BackColor = vbRed              '&HFF
lblColor(14).BackColor = vbMagenta          '&HFF00FF
'******************load the Grid**********************
LoadGrid 32

'***************Default Mouse Pointer for Paint**********
LargePicture.MousePointer = 99

End Sub


Private Sub LargePicture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'*********to change backGround Color***************
If LargePicture.MousePointer = 10 Then      '10=vbUpArrow
      
        b = LargePicture.Point(X, Y)
        If b = MyColor Then Exit Sub
    For j = 0 To SmallPicture.ScaleWidth - 1
        For p = 0 To SmallPicture.ScaleHeight - 1
            c = SmallPicture.Point(j, p)
            If c = b Then SmallPicture.PSet (j, p), MyColor
        Next p
    Next j
   
 
  Set Picture0 = SmallPicture.Image

 '******updating LargePicture from SmallPicture************
  LargePicture.PaintPicture Picture0, 0, 0, 321, 321
    LoadGrid 32
   LargePicture.MousePointer = 99
     Exit Sub
End If

'*****************Normal Point(Paint) draw*******************
If Button = vbLeftButton And LargePicture.MousePointer = 99 Then
X1 = 0:  Y1 = 0
  For j = 0 To 31
      For p = 0 To 31
          If X < (X1 + ValAdd) And X > X1 And Y < (Y1 + ValAdd) And Y > Y1 Then
             LargePicture.Line (X1 + 1, Y1 + 1)-(X1 + upPoint, Y1 + upPoint), MyColor, BF
             SmallPicture.PSet (X1 \ ValAdd, Y1 \ ValAdd), MyColor
          End If
          X1 = X1 + ValAdd
          If X1 = 320 Then
             X1 = 0
             Y1 = Y1 + ValAdd
          End If
      Next p
  Next j
  Exit Sub
  End If
End Sub

Private Sub LargePicture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'****************Draw Line********************
If Button = vbLeftButton Then
X1 = 0:  Y1 = 0
  For j = 0 To 31
      For p = 0 To 31
          If X < X1 + ValAdd And X > X1 And Y < Y1 + ValAdd And Y > Y1 Then
             LargePicture.Line (X1 + 1, Y1 + 1)-(X1 + upPoint, Y1 + upPoint), MyColor, BF
             SmallPicture.PSet (X1 \ ValAdd, Y1 \ ValAdd), MyColor
          End If
          X1 = X1 + ValAdd
          If X1 = 320 Then
             X1 = 0
             Y1 = Y1 + ValAdd
          End If
      Next p
  Next j
  
  End If
  
End Sub

Private Sub lblColor_Click(Index As Integer)
MyColor = lblColor(Index).BackColor
Pallet.BackColor = lblColor(Index).BackColor
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    
    Case "New":    New_Icon
      
    Case "Paint":  Toolbar1.Buttons("Change").Value = tbrUnpressed
      Toolbar1.Buttons("Paint").Value = tbrPressed
      LargePicture.MousePointer = 99

    Case "Change": Toolbar1.Buttons("Paint").Value = tbrUnpressed
      Toolbar1.Buttons("Change").Value = tbrPressed
      LargePicture.MousePointer = 10
    
    Case "Up":     movePict "up"
    Case "Down":   movePict "down"
    Case "Left":   movePict "left"
    Case "Right":  movePict "right"

Case "flipHorizontal"
Set Picture0 = SmallPicture.Image
LargePicture.PaintPicture Picture0, LargePicture.ScaleWidth - 1, 0, -LargePicture.ScaleWidth, LargePicture.ScaleHeight
SmallPicture.PaintPicture LargePicture.Image, 0, 0, 32, 32
        LoadGrid GridFormat
        
Case "FlipVert"
Set Picture0 = SmallPicture.Image
LargePicture.PaintPicture Picture0, 0, LargePicture.ScaleHeight - 1, LargePicture.ScaleWidth, -LargePicture.ScaleHeight
SmallPicture.PaintPicture LargePicture.Image, 0, 0, 32, 32
        LoadGrid GridFormat
'
'
    Case "Rotate": Rotat_Icon
    Case "Clear"
LargePicture.Picture = LoadPicture
SmallPicture.Picture = LoadPicture
LoadGrid GridFormat

    Case "Exit":   Unload Me
    Case "Extract"
    
    '*********
    CommonDialog1.Filter = "Executable (*.exe)|*.exe|DLL(*.dll)|*.dll"
    CommonDialog1.ShowOpen
    
    If CommonDialog1.filename <> "" Then
        TotalApps = TotalApps + 1
        ReDim MyApp(TotalApps)
        MyApp(TotalApps) = CommonDialog1.filename
        GetAppIcon (CommonDialog1.filename)
        LoadGrid 32
    End If
'*****************
  End Select

End Sub
Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
If ButtonMenu.Parent.Index = 4 Then
If ButtonMenu.Index = 1 Then save_Ico
If ButtonMenu.Index = 2 Then save_BMP
End If
If ButtonMenu.Parent.Index = 2 Then
If ButtonMenu.Index = 1 Then Open_Icon
If ButtonMenu.Index = 2 Then Open_BMP
End If

End Sub

Private Sub Open_Icon()
  Dim MyFile As String
  CommonDialog1.CancelError = False
  
'==================
'cdlOFNFileMustExist Definition:
'==================
  'Specifies that the user can enter only
  'names of existing files in the File Name text box.
  'If this flag is set and the user enters an invalid filename,
  'MyColor warning is displayed.
  'This flag automatically sets the [cdlOFNPathMustExist] flag
 
 CommonDialog1.CancelError = False
  CommonDialog1.Flags = cdlOFNFileMustExist
  CommonDialog1.Filter = "Icons (*.ico)|*.ico"
  CommonDialog1.ShowOpen
  MyFile = CommonDialog1.filename

If MyFile <> "" Then
  lblStatusBar.Caption = " " & MyFile & "... " & FileLen(MyFile) & " byte"
   
  SmallPicture = LoadPicture(MyFile)
  LargePicture.PaintPicture SmallPicture.Image, 0, 0, 321, 321
  LoadGrid 32
Else
  Exit Sub
End If

End Sub
Public Sub Open_BMP()
Dim MyFile As String
  
 
 CommonDialog1.CancelError = False
  CommonDialog1.Flags = cdlOFNFileMustExist
  CommonDialog1.Filter = "Bitmap(*.bmp)|*.bmp"
  CommonDialog1.ShowOpen
  MyFile = CommonDialog1.filename

If MyFile <> "" Then
  lblStatusBar.Caption = " " & MyFile & "... " & FileLen(MyFile) & " byte"
   Image1.Picture = LoadPicture(MyFile)
  LargePicture.PaintPicture Image1.Picture, 0, 0, 321, 321
  
  SmallPicture.PaintPicture LargePicture.Image, 0, 0, 32, 32
  LoadGrid 32
Else
  Exit Sub
End If

End Sub

Private Sub LoadGrid(GridFormat As Integer, Optional backClr As Long)

GridFormat = 32
upPoint = 9
ValAdd = 10
GridStep = 10

If backClr = 0 Then backClr = vbBlack

For F = 0 To LargePicture.ScaleHeight Step GridStep
LargePicture.Line (0, F)-(LargePicture.ScaleWidth, F), backClr
Next F
For F = 0 To LargePicture.ScaleWidth Step GridStep
LargePicture.Line (F, 0)-(F, LargePicture.ScaleHeight), backClr
Next F
End Sub

Public Sub New_Icon()
SmallPicture.Picture = LoadPicture("")
LargePicture = LoadPicture
LoadGrid 32
End Sub

Private Sub save_Ico()
Dim MyFile As String

CommonDialog1.CancelError = False
CommonDialog1.filename = ""
CommonDialog1.Flags = cdlOFNOverwritePrompt + cdlOFNNoReadOnlyReturn
CommonDialog1.Filter = "Icons (*.ico)|*.ico"
CommonDialog1.ShowSave
CommonDialog1.FilterIndex = 1
MyFile = CommonDialog1.filename
If MyFile <> "" Then

Set Picture0 = SmallPicture.Image

Dim li As ListImage
Dim ThePic  As StdPicture

Set li = ImageList1.ListImages.Add(, , Picture0)
Set ThePic = li.ExtractIcon
SavePicture ThePic, MyFile
End If
End Sub
Public Sub save_BMP()
Dim MyFile As String

CommonDialog1.CancelError = False
CommonDialog1.filename = ""
CommonDialog1.Flags = cdlOFNOverwritePrompt + cdlOFNNoReadOnlyReturn
CommonDialog1.Filter = "Bitmaps (*.bmp)|*.bmp"
CommonDialog1.ShowSave
CommonDialog1.FilterIndex = 1
MyFile = CommonDialog1.filename
If MyFile <> "" Then

Set Picture0 = LargePicture.Image
SavePicture Picture0, MyFile   'this will save it as bmp
End If
End Sub


Private Sub movePict(moveType As String)
  Select Case moveType
    
    Case "left":    RtoL -10, 0
    Case "right":   RtoL 10, 0
    Case "up":      RtoL 0, -10
    Case "down":    RtoL 0, 10
  End Select
  
 LoadGrid 32
End Sub

Private Sub RtoL(Pr1 As Long, Pr2 As Long)

 Set Picture0 = SmallPicture.Image
 
 LargePicture.PaintPicture Picture0, Pr1, Pr2, 319, 319
SmallPicture.PaintPicture LargePicture.Image, 0, 0, 32, 32
   
  End Sub

Public Sub Rotat_Icon()
Dim MyPic As PictureBox
Set MyPic = SmallPicture

Dim p As Long
Dim j As Long

 For j = 0 To 31
      For p = 0 To 31
          SmallPicture.PSet (j, p), MyPic.Point(p, j)
      Next p
  Next j

Set Picture0 = SmallPicture.Image
    SmallPicture.Picture = LoadPicture("")
    SmallPicture.PaintPicture Picture0, 31, 0, -32, 32

  LargePicture.PaintPicture SmallPicture.Image, 0, 0, 321, 321
  LoadGrid 32

  SmallPicture.PaintPicture LargePicture.Image, 0, 0, 32, 32
Set Picture0 = SmallPicture.Image

End Sub


Sub GetAppIcon(ProgramPath As String)

    Dim lngIcon As Long
    Dim strPath As String
    
    strPath = ProgramPath
    
    LargePicture.Picture = LoadPicture("")
    lngIcon = ExtractAssociatedIcon(App.hInstance, strPath, 1)
    DrawIconEx LargePicture.hdc, 0, 0, lngIcon, 321, 321, 0, 0, DI_NORMAL
    DestroyIcon lngIcon
    SmallPicture.PaintPicture LargePicture.Image, 0, 0, 32, 32
End Sub
Private Sub Form_Activate()
Dim K As Long, j As Long, L As Long
Dim i As Long
j = 0
L = 0
i = 0
For K = 0 To 255 Step 15
For j = 0 To 255 Step 20
For L = 0 To 255 Step 20
Pic.Line (0, i)-Step(Pic.ScaleWidth, i), RGB(K, j, L), B
i = i + 1
Next L
i = i + 1
Next j
i = i + 1
Next K
Pic.Refresh

End Sub
Private Sub Pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Pallet2.BackColor = Pic.Point(X, Y)
MyColor = Pallet2.BackColor
End Sub

