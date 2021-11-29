VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "*\AAxMCCombo2.vbp"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "AxioUK MultiColumn/MultiLine ComboBox"
   ClientHeight    =   5100
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9750
   FillColor       =   &H00C07000&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   340
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   650
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3705
      TabIndex        =   49
      Text            =   "0"
      Top             =   1740
      Width           =   420
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3705
      TabIndex        =   48
      Text            =   "1"
      Top             =   1395
      Width           =   420
   End
   Begin axMCCombo2.axMCCombo axMCCombo1 
      Height          =   420
      Left            =   2730
      TabIndex        =   45
      Top             =   2595
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   741
      HeaderH         =   24
      LineColor       =   15790320
      GridStyle       =   3
      Striped         =   -1  'True
      StripedColor    =   16645629
      SelColor        =   -2147483635
      ItemH           =   0
      BorderColor     =   8388608
      BorderWidth     =   1
      CornerCurve     =   1
      Header          =   -1  'True
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisibleRows     =   8
      DropWidth       =   0
      ButtonColorPress=   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IcoFont"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconCharCode    =   60007
      IconForeColor   =   0
      IcoPaddingX     =   5
      IcoPaddingY     =   4
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2850
      Left            =   5925
      TabIndex        =   44
      Top             =   2025
      Width           =   3570
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   285
         TabIndex        =   47
         Text            =   "Combo Standar VB6"
         Top             =   1920
         Width           =   2205
      End
      Begin axMCCombo2.axMCCombo axMCCombo2 
         Height          =   750
         Left            =   255
         TabIndex        =   46
         Top             =   450
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   1323
         HeaderH         =   24
         LineColor       =   15790320
         GridStyle       =   3
         Striped         =   -1  'True
         StripedColor    =   16645629
         SelColor        =   -2147483635
         ItemH           =   0
         BorderColor     =   9471874
         BorderWidth     =   1
         CornerCurve     =   1
         Header          =   -1  'True
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisibleRows     =   8
         DropWidth       =   0
         MultiLine       =   -1  'True
         ButtonColorPress=   0
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IcoFont"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconCharCode    =   61357
         IconForeColor   =   0
         IcoPaddingX     =   -5
         IcoPaddingY     =   0
      End
   End
   Begin VB.CheckBox Check3 
      Alignment       =   1  'Right Justify
      Caption         =   "Multiline ?"
      Height          =   195
      Left            =   5970
      TabIndex        =   43
      Top             =   1650
      Value           =   1  'Checked
      Width           =   1260
   End
   Begin VB.PictureBox Color 
      Height          =   285
      Index           =   7
      Left            =   1575
      ScaleHeight     =   225
      ScaleWidth      =   570
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   3540
      Width           =   630
   End
   Begin VB.CommandButton cmdCharCode 
      Caption         =   "Set"
      Height          =   360
      Left            =   5115
      TabIndex        =   39
      Top             =   960
      Width           =   480
   End
   Begin VB.TextBox txtCharCode 
      Height          =   330
      Left            =   4290
      TabIndex        =   37
      Text            =   "&Hea67"
      Top             =   960
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   360
      Left            =   10575
      TabIndex        =   36
      Top             =   435
      Width           =   990
   End
   Begin VB.PictureBox Picture1 
      Height          =   300
      Left            =   10110
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   255
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   450
      Width           =   315
   End
   Begin VB.TextBox txt1 
      Height          =   330
      Left            =   3345
      TabIndex        =   30
      Top             =   3435
      Width           =   1965
   End
   Begin VB.TextBox txt2 
      Height          =   330
      Left            =   3345
      TabIndex        =   29
      Top             =   3780
      Width           =   1965
   End
   Begin VB.TextBox txt3 
      Height          =   330
      Left            =   3345
      TabIndex        =   28
      Top             =   4125
      Width           =   1965
   End
   Begin VB.TextBox txt4 
      Height          =   330
      Left            =   3345
      TabIndex        =   27
      Top             =   4470
      Width           =   1965
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1785
      TabIndex        =   25
      Text            =   "8"
      Top             =   1110
      Width           =   420
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   1005
      TabIndex        =   24
      Top             =   225
      Width           =   1200
   End
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      Caption         =   "StripedList"
      Height          =   195
      Left            =   2760
      TabIndex        =   23
      Top             =   1050
      Value           =   1  'Checked
      Width           =   1260
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1680
      TabIndex        =   21
      Text            =   "0"
      Top             =   1785
      Width           =   525
   End
   Begin VB.PictureBox Color 
      Height          =   285
      Index           =   6
      Left            =   1575
      ScaleHeight     =   225
      ScaleWidth      =   570
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4620
      Width           =   630
   End
   Begin VB.PictureBox Color 
      Height          =   285
      Index           =   5
      Left            =   1575
      ScaleHeight     =   225
      ScaleWidth      =   570
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4260
      Width           =   630
   End
   Begin VB.PictureBox Color 
      Height          =   285
      Index           =   4
      Left            =   1575
      ScaleHeight     =   225
      ScaleWidth      =   570
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3900
      Width           =   630
   End
   Begin VB.PictureBox Color 
      Height          =   285
      Index           =   3
      Left            =   1575
      ScaleHeight     =   225
      ScaleWidth      =   570
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3195
      Width           =   630
   End
   Begin VB.PictureBox Color 
      Height          =   285
      Index           =   2
      Left            =   1575
      ScaleHeight     =   225
      ScaleWidth      =   570
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2850
      Width           =   630
   End
   Begin VB.PictureBox Color 
      Height          =   285
      Index           =   1
      Left            =   1575
      ScaleHeight     =   225
      ScaleWidth      =   570
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2490
      Width           =   630
   End
   Begin VB.PictureBox Color 
      Height          =   285
      Index           =   0
      Left            =   1575
      ScaleHeight     =   225
      ScaleWidth      =   570
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2130
      Width           =   630
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   45
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton Option1 
      Caption         =   "JList"
      Height          =   240
      Index           =   1
      Left            =   2760
      TabIndex        =   4
      Top             =   435
      Width           =   1305
   End
   Begin VB.OptionButton Option1 
      Caption         =   "JCombo"
      Height          =   240
      Index           =   0
      Left            =   2760
      TabIndex        =   3
      Top             =   165
      Value           =   -1  'True
      Width           =   1305
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Header List"
      Height          =   195
      Left            =   2760
      TabIndex        =   0
      Top             =   795
      Value           =   1  'Checked
      Width           =   1260
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1785
      TabIndex        =   2
      Text            =   "0"
      Top             =   1455
      Width           =   420
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CornerRound"
      Height          =   195
      Left            =   2580
      TabIndex        =   51
      Top             =   1800
      Width           =   960
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BorderWidth"
      Height          =   195
      Left            =   2625
      TabIndex        =   50
      Top             =   1455
      Width           =   900
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GridLines"
      Height          =   195
      Left            =   1035
      TabIndex        =   42
      Top             =   45
      Width           =   645
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IconForeColor"
      Height          =   195
      Left            =   480
      TabIndex        =   41
      Top             =   3585
      Width           =   1020
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Debe seleccionar el IconFont            IcoFont IconChar Code"
      Height          =   795
      Left            =   4305
      TabIndex        =   38
      Top             =   150
      Width           =   1290
   End
   Begin VB.Line Line1 
      X1              =   384
      X2              =   384
      Y1              =   11
      Y2              =   327
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Columna4"
      Height          =   195
      Left            =   2370
      TabIndex        =   34
      Top             =   4545
      Width           =   870
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Columna3"
      Height          =   195
      Left            =   2370
      TabIndex        =   33
      Top             =   4200
      Width           =   870
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Columna2"
      Height          =   195
      Left            =   2370
      TabIndex        =   32
      Top             =   3840
      Width           =   870
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Columna1"
      Height          =   195
      Left            =   2370
      TabIndex        =   31
      Top             =   3495
      Width           =   870
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Visible Rows"
      Height          =   195
      Left            =   735
      TabIndex        =   26
      Top             =   1170
      Width           =   870
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DropWidth (0=Auto)"
      Height          =   195
      Left            =   180
      TabIndex        =   22
      Top             =   1845
      Width           =   1485
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "StripBackColor"
      Height          =   195
      Left            =   480
      TabIndex        =   14
      Top             =   4665
      Width           =   1020
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ForeColor"
      Height          =   195
      Left            =   795
      TabIndex        =   13
      Top             =   3240
      Width           =   705
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GridLineColor"
      Height          =   195
      Left            =   555
      TabIndex        =   12
      Top             =   3945
      Width           =   945
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SelectionColor"
      Height          =   195
      Left            =   480
      TabIndex        =   11
      Top             =   4305
      Width           =   1020
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ButtonColorPress"
      Height          =   195
      Left            =   255
      TabIndex        =   10
      Top             =   2895
      Width           =   1245
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BorderColor"
      Height          =   195
      Left            =   795
      TabIndex        =   9
      Top             =   2535
      Width           =   705
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BackColor"
      Height          =   195
      Left            =   795
      TabIndex        =   8
      Top             =   2175
      Width           =   705
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0312
      Height          =   975
      Left            =   6555
      TabIndex        =   6
      Top             =   405
      Width           =   2610
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AxMCCombo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   7005
      TabIndex        =   5
      Top             =   90
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ColumnInBox"
      Height          =   195
      Left            =   660
      TabIndex        =   1
      Top             =   1515
      Width           =   945
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IsVisible As Boolean

'Create a new project, add a command button and a picture box to the project, load a picture into the picture box.
'Paste this code into Form1
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Dim PicBits() As Byte, PicInfo As BITMAP, Cnt As Long

Private Sub Check3_Click()
axMCCombo1.MultiLine = Check3.Value
End Sub

Private Sub cmdCharCode_Click()
On Error Resume Next
axMCCombo1.IconCharCode = txtCharCode.Text
axMCCombo2.IconCharCode = txtCharCode.Text
End Sub

Private Sub Command1_Click()
    'KPD-Team 1999
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    'Get information (such as height and width) about the picturebox
    GetObject Picture1.Image, Len(PicInfo), PicInfo
    'reallocate storage space
    ReDim PicBits(1 To PicInfo.bmWidth * PicInfo.bmHeight * 3) As Byte
    'Copy the bitmapbits to the array
    GetBitmapBits Picture1.Image, UBound(PicBits), PicBits(1)
    'Invert the bits
    For Cnt = 1 To UBound(PicBits)
        PicBits(Cnt) = 255 - PicBits(Cnt)
    Next Cnt
    'Set the bits back to the picture
    SetBitmapBits Picture1.Image, UBound(PicBits), PicBits(1)
    'refresh
    Picture1.Refresh
End Sub



Private Sub axMCCombo1_ItemClick(Item As Long)
    txt1 = axMCCombo1.ItemText(Item, 0)
    txt2 = axMCCombo1.ItemText(Item, 1)
    txt3 = axMCCombo1.ItemText(Item, 2)

    txt1.SelStart = 0
    txt1.SelLength = Len(txt2)
End Sub

Private Sub Check1_Click()
axMCCombo1.Header = Check1.Value
End Sub

Private Sub Check2_Click()
axMCCombo1.StripedGrid = Check2.Value
End Sub

Private Sub Color_Click(Index As Integer)
cDialog.DialogTitle = "Select Color"
cDialog.ShowColor
Color(Index).BackColor = cDialog.Color

With axMCCombo1
        .BackColor = Color(0).BackColor
        .BorderColor = Color(1).BackColor
        .ButtonColorPress = Color(2).BackColor
        .ForeColor = Color(3).BackColor
        .GridLineColor = Color(4).BackColor
        .SelectionColor = Color(5).BackColor
        .StripBackColor = Color(6).BackColor
        .IconForeColor = Color(7).BackColor
End With

End Sub

Private Sub Form_Load()
'gbAllowSubclassing = True
'SubclassToSeeMessages Me.hWnd

IsVisible = False

Dim i As Long

    With axMCCombo1
        .AddColumn "ControlName"
        .AddColumn "Creator"
        .AddColumn "Points"

        For i = 1 To 20
            .AddItem "axJColCombo_" & i, 0
            .ItemText(.ItemCount - 1, 1) = "AxioUK_" & i
            .ItemText(.ItemCount - 1, 2) = 25 + i
        Next
        .ColWidthAutoSize
        .ColumnInBox = CInt(Text4.Text)
        
        Color(0).BackColor = .BackColor
        Color(1).BackColor = .BorderColor
        Color(2).BackColor = .ButtonColorPress
        Color(3).BackColor = .ForeColor
        Color(4).BackColor = .GridLineColor
        Color(5).BackColor = .SelectionColor
        Color(6).BackColor = .StripBackColor
        Color(7).BackColor = .IconForeColor
    End With
                
    With axMCCombo2
          .AddColumn "ControlName"
          .AddColumn "Creator"
          .AddColumn "Points"
  
          For i = 1 To 20
              .AddItem "axJColCombo_" & i, 0
              .ItemText(.ItemCount - 1, 1) = "AxioUK_" & i
              .ItemText(.ItemCount - 1, 2) = 25 + i
          Next
          .ColWidthAutoSize
          .ColumnInBox = CInt(Text4.Text)
    End With
                
List1.AddItem "0 - None", 0
List1.AddItem "1 - Horizontal", 1
List1.AddItem "2 - Vertical", 2
List1.AddItem "3 - Both", 3

'List2.AddItem "Up", 0
'List2.AddItem "Down", 1
'List2.AddItem "Left", 2
'List2.AddItem "Right", 3

End Sub


Private Sub List1_Click()
axMCCombo1.GridLineStyle = List1.ListIndex
axMCCombo2.GridLineStyle = List1.ListIndex

End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
  Case 0
    axMCCombo1.ComboStyle = 0
    axMCCombo2.ComboStyle = 0
  Case 1
    axMCCombo1.ComboStyle = 1
    axMCCombo2.ComboStyle = 1
End Select
End Sub

Private Sub Text1_Change()
axMCCombo1.DropWidth = Text1.Text
End Sub

Private Sub Text2_Change()
axMCCombo1.VisibleRows = Text2.Text
End Sub

Private Sub Text3_Change()
On Error Resume Next
axMCCombo1.BorderWidth = CInt(Text3.Text)
axMCCombo2.BorderWidth = CInt(Text3.Text)
End Sub

Private Sub Text4_Change()
On Error Resume Next
axMCCombo1.ColumnInBox = CInt(Text4.Text)
axMCCombo2.ColumnInBox = CInt(Text4.Text)
End Sub

Private Sub Text5_Change()
On Error Resume Next
axMCCombo1.CornerRound = CInt(Text5.Text)
axMCCombo2.CornerRound = CInt(Text5.Text)

End Sub
