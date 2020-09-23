VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmGradient 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gradient"
   ClientHeight    =   2508
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   9192
   Icon            =   "frmGradient.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2508
   ScaleWidth      =   9192
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picExample 
      AutoRedraw      =   -1  'True
      Height          =   1092
      Left            =   5160
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   87
      TabIndex        =   13
      ToolTipText     =   "Drag to the left box to add to custom"
      Top             =   120
      Width           =   1092
      Begin VB.Line lnAngle 
         Visible         =   0   'False
         X1              =   40
         X2              =   40
         Y1              =   40
         Y2              =   60
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   40
      Left            =   6360
      ScaleHeight     =   3
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   195
      TabIndex        =   32
      Top             =   1185
      Width           =   2345
      Begin VB.Line Line6 
         BorderColor     =   &H80000003&
         X1              =   0
         X2              =   1
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   195
         Y1              =   1
         Y2              =   1
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   372
      Left            =   6360
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   196
      TabIndex        =   31
      Top             =   -250
      Width           =   2350
      Begin VB.Line Line4 
         BorderColor     =   &H80000003&
         X1              =   0
         X2              =   0
         Y1              =   29
         Y2              =   31
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000003&
         X1              =   0
         X2              =   200
         Y1              =   29
         Y2              =   29
      End
      Begin VB.Line Line1 
         X1              =   1
         X2              =   200
         Y1              =   30
         Y2              =   30
      End
   End
   Begin VB.Frame fraCustom 
      Height          =   1212
      Left            =   6360
      TabIndex        =   26
      Top             =   0
      Width           =   2532
      Begin VB.PictureBox picCustom 
         BackColor       =   &H00FFFFFF&
         Height          =   1112
         Left            =   0
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   1068
         ScaleWidth      =   2304
         TabIndex        =   28
         Top             =   100
         Width           =   2352
         Begin VB.PictureBox picIcon 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   0
            Left            =   480
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   20
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   20
            TabIndex        =   30
            Top             =   360
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picIconHolder 
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   312
            Left            =   444
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   312
            ScaleWidth      =   312
            TabIndex        =   29
            Top             =   324
            Visible         =   0   'False
            Width           =   312
         End
      End
      Begin VB.VScrollBar vscrCustom 
         Enabled         =   0   'False
         Height          =   1112
         LargeChange     =   354
         Left            =   2340
         SmallChange     =   14
         TabIndex        =   27
         Top             =   100
         Width           =   175
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   432
      Left            =   7680
      TabIndex        =   25
      Top             =   1980
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   432
      Left            =   7680
      TabIndex        =   24
      Top             =   1440
      Width           =   1212
   End
   Begin VB.CommandButton cdmOpen 
      Caption         =   "Open"
      Height          =   272
      Left            =   6480
      TabIndex        =   23
      Top             =   2140
      Width           =   1092
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   272
      Left            =   6480
      TabIndex        =   22
      Top             =   1800
      Width           =   1092
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   272
      Left            =   6480
      TabIndex        =   21
      Top             =   1440
      Width           =   1092
   End
   Begin MSComDlg.CommonDialog cdgMain 
      Left            =   6480
      Top             =   2520
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Frame2 
      Height          =   1092
      Left            =   5160
      TabIndex        =   16
      Top             =   1320
      Width           =   1212
      Begin VB.TextBox txtDeg 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   144
         Left            =   996
         TabIndex        =   36
         Text            =   "°"
         Top             =   492
         Width           =   80
      End
      Begin VB.TextBox txtAngle 
         Appearance      =   0  'Flat
         Height          =   192
         Left            =   720
         MaxLength       =   3
         TabIndex        =   35
         Text            =   "45"
         Top             =   480
         Width           =   372
      End
      Begin VB.OptionButton optRadial 
         Caption         =   "Radial"
         Height          =   252
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   852
      End
      Begin VB.OptionButton optLinear 
         Caption         =   "Linear"
         Height          =   252
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   852
      End
      Begin VB.Line Line8 
         X1              =   480
         X2              =   600
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line7 
         X1              =   480
         X2              =   480
         Y1              =   480
         Y2              =   600
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1092
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Width           =   4952
      Begin VB.TextBox txtColorHex 
         Alignment       =   2  'Center
         Height          =   216
         Index           =   0
         Left            =   3240
         TabIndex        =   4
         Text            =   "FF"
         Top             =   240
         Width           =   372
      End
      Begin VB.TextBox txtColorDec 
         Alignment       =   2  'Center
         Height          =   216
         Index           =   0
         Left            =   2796
         TabIndex        =   1
         Text            =   "255"
         Top             =   240
         Width           =   372
      End
      Begin VB.HScrollBar hscrColor 
         Height          =   180
         Index           =   2
         LargeChange     =   5
         Left            =   360
         Max             =   255
         TabIndex        =   9
         Top             =   744
         Value           =   255
         Width           =   2352
      End
      Begin VB.HScrollBar hscrColor 
         Height          =   180
         Index           =   1
         LargeChange     =   5
         Left            =   360
         Max             =   255
         TabIndex        =   8
         Top             =   504
         Value           =   255
         Width           =   2352
      End
      Begin VB.HScrollBar hscrColor 
         Height          =   180
         Index           =   0
         LargeChange     =   5
         Left            =   360
         Max             =   255
         TabIndex        =   7
         Top             =   264
         Value           =   255
         Width           =   2352
      End
      Begin VB.PictureBox PicColorPicker 
         BackColor       =   &H00FFFFFF&
         Height          =   732
         Left            =   3720
         MouseIcon       =   "frmGradient.frx":038A
         MousePointer    =   99  'Custom
         ScaleHeight     =   684
         ScaleWidth      =   1068
         TabIndex        =   15
         Top             =   240
         Width           =   1112
      End
      Begin VB.TextBox txtColorDec 
         Alignment       =   2  'Center
         Height          =   216
         Index           =   1
         Left            =   2796
         TabIndex        =   2
         Text            =   "255"
         Top             =   480
         Width           =   372
      End
      Begin VB.TextBox txtColorDec 
         Alignment       =   2  'Center
         Height          =   216
         Index           =   2
         Left            =   2796
         TabIndex        =   3
         Text            =   "255"
         Top             =   720
         Width           =   372
      End
      Begin VB.TextBox txtColorHex 
         Alignment       =   2  'Center
         Height          =   216
         Index           =   1
         Left            =   3240
         TabIndex        =   5
         Text            =   "FF"
         Top             =   480
         Width           =   372
      End
      Begin VB.TextBox txtColorHex 
         Alignment       =   2  'Center
         Height          =   216
         Index           =   2
         Left            =   3240
         TabIndex        =   6
         Text            =   "FF"
         Top             =   720
         Width           =   372
      End
      Begin VB.Label Label3 
         Caption         =   "B :"
         Height          =   252
         Left            =   120
         TabIndex        =   19
         Top             =   718
         Width           =   252
      End
      Begin VB.Label Label2 
         Caption         =   "G :"
         Height          =   252
         Left            =   110
         TabIndex        =   18
         Top             =   478
         Width           =   252
      End
      Begin VB.Label Label1 
         Caption         =   "R :"
         Height          =   252
         Left            =   120
         TabIndex        =   17
         Top             =   238
         Width           =   252
      End
   End
   Begin VB.Frame fraPointer 
      Height          =   1212
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4952
      Begin VB.PictureBox picPointer 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   0
         Left            =   120
         Picture         =   "frmGradient.frx":0694
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   11
         TabIndex        =   34
         Top             =   168
         Visible         =   0   'False
         Width           =   130
      End
      Begin VB.ComboBox cboPointers 
         Height          =   288
         Left            =   4500
         Style           =   1  'Simple Combo
         TabIndex        =   20
         Text            =   "cboPointers"
         Top             =   516
         Width           =   312
      End
      Begin VB.PictureBox picGradient 
         AutoRedraw      =   -1  'True
         Height          =   252
         Left            =   340
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   334
         TabIndex        =   12
         Top             =   840
         Width           =   4052
      End
      Begin MSComctlLib.Toolbar tbrSnap 
         Height          =   312
         Left            =   4524
         TabIndex        =   33
         Top             =   120
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   550
         ButtonWidth     =   487
         ButtonHeight    =   466
         ImageList       =   "imglstSnap"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Snap"
               Object.ToolTipText     =   "Snap"
               ImageIndex      =   1
               Style           =   1
            EndProperty
         EndProperty
      End
      Begin VB.Line lnSnapPointers 
         Index           =   0
         X1              =   360
         X2              =   360
         Y1              =   600
         Y2              =   720
      End
      Begin VB.Line lnPointer 
         BorderColor     =   &H80000015&
         X1              =   360
         X2              =   4360
         Y1              =   492
         Y2              =   492
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   360
         X2              =   4360
         Y1              =   504
         Y2              =   504
      End
   End
   Begin MSComctlLib.ImageList imglstGradientPics 
      Left            =   4680
      Top             =   2520
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imglstSnap 
      Left            =   5880
      Top             =   2520
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGradient.frx":09CA
            Key             =   "Snap"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstPointers 
      Left            =   5280
      Top             =   2520
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   11
      ImageHeight     =   21
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGradient.frx":0D1E
            Key             =   "UnSelected"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGradient.frx":1066
            Key             =   "Selected"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPointer 
      Caption         =   "Pointer"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
   End
End
Attribute VB_Name = "frmGradient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bolLockPointer As Boolean
Public intSelPointer As Integer
Public bolSkipExample As Boolean
Public intSnapCount As Integer
Private SnapPoints() As Long
Public intSelCustom As Integer
Public bolCustomFocus As Boolean

Private Sub PaintPointer(NewColor As Long, Index As Integer, Selected As Boolean, Optional DontSetColor As Boolean)
Dim X, Y As Integer
Dim lngColor As Long

    If Selected = False Then
        picPointer(Index).Picture = imglstPointers.ListImages("UnSelected").Picture
    Else
        picPointer(Index).Picture = imglstPointers.ListImages("Selected").Picture
    End If
    For Y = 2 To 15
        picPointer(Index).Line (2, Y)-(9, Y), NewColor
    Next
    For X = 3 To 7
        For Y = 16 To 18
            lngColor = picPointer(Index).Point(X, Y)
            If lngColor = RGB(255, 0, 255) Then lngColor = NewColor
            picPointer(Index).PSet (X, Y), lngColor
        Next
    Next
    If DontSetColor = False Then
        PointColors(picPointer(Index).Tag).Color = NewColor
    End If
End Sub

Private Sub cboPointers_Change()
    intSnapCount = Val(cboPointers.Text) - 1
    DrawSnapPointers intSnapCount
End Sub

Private Sub cboPointers_GotFocus()
    cboPointers.SelStart = 0
    cboPointers.SelLength = Len(cboPointers.Text)
End Sub

Private Sub cboPointers_KeyPress(KeyAscii As Integer)
'alleen getallen
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub cdmOpen_Click()
Dim Temp() As GradientFiles
On Error GoTo Errorhandler
    
    cdgMain.Filter = "Custom Gradient Files(*.cgf)|*.cgf|All Files(*.*)|*.*"
    cdgMain.ShowOpen
    LoadGradient cdgMain.FileName, Temp()
    
    LoadCustom Temp
    
Exit Sub
Errorhandler:
    If Err.Number <> 32755 Then 'Cancel was selected.
        
    End If
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNew_Click()
    NewCustom
End Sub

Private Sub cmdOK_Click()
    'place here your own picturebox
    'EXAMPLE:
    'PaintGradient frmMain.picMain, PointColors, GradientType
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo Errorhandler
    
    cdgMain.Filter = "Custom Gradient Files(*.cgf)|*.cgf"
    cdgMain.ShowSave
    SaveGradient cdgMain.FileName, CustomGradient()
    
Exit Sub
Errorhandler:
    If Err.Number <> 32755 Then 'Cancel was selected.
        
    End If
End Sub

Private Sub Form_Load()
Dim intCount As Integer
    
    ' Get DataBase Workspace.
    Set objWorkSpace = DBEngine.Workspaces(0)
    
    For intCount = 1 To 25
        cboPointers.AddItem intCount
        Load lnSnapPointers(intCount)
    Next
    intSnapCount = 10
    cboPointers.ListIndex = intSnapCount - 1
    DrawSnapPointers intSnapCount
    
    NewPointer &H0, 0, True
    intSelPointer = 1
    NewPointer &HFF&, 100
    NewPointer 0
    PaintExamples
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("Do you want to vote for this code?", vbQuestion + vbYesNo, "Vote ?") = vbYes Then
        ShellExecute hWnd, "open", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=33973&lngWId=1", vbNullChar, vbNullChar, SW_NORMAL
    End If
End Sub

Private Sub hscrColor_Change(Index As Integer)
    PicColorPicker.BackColor = RGB(hscrColor(0).Value, _
                             hscrColor(1).Value, _
                             hscrColor(2).Value)
    txtColorDec(Index).Text = hscrColor(Index).Value
    txtColorHex(Index).Text = Hex(hscrColor(Index).Value)
    If bolSkipExample <> True Then
        PaintPointer PicColorPicker.BackColor, intSelPointer, True
        PaintExamples
    End If
    If bolCustomFocus = False Then
        picIconHolder.Visible = False
    End If
End Sub


Private Sub mnuDelete_Click()
    If bolCustomFocus = False Then
        DelPointer intSelPointer
    Else
        DelCustom intSelCustom
    End If
End Sub

Private Sub optLinear_Click()
    GradientType = Linear
    PaintExamples
End Sub

Private Sub optRadial_Click()
    GradientType = Radial
    PaintExamples
End Sub

Private Sub PicColorPicker_DblClick()
On Error GoTo Errorhandler

    cdgMain.Color = PicColorPicker.BackColor
    cdgMain.ShowColor
    hscrColor(0).Value = DefineRGB(cdgMain.Color).Red
    hscrColor(1).Value = DefineRGB(cdgMain.Color).Green
    hscrColor(2).Value = DefineRGB(cdgMain.Color).Blue
Exit Sub
Errorhandler:
    If Err.Number <> 32755 Then 'Cancel was selected.
        
    End If
End Sub

Private Sub PicColorPicker_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetCapture(PicColorPicker.hWnd)
End Sub
Private Sub PicColorPicker_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngDC, lngResult, lnghWnd As Long
Dim MousePoint As POINTAPI
Dim Colors As RGBColor
    bolSkipExample = True
    Me.MousePointer = 99
    GetCursorPos MousePoint
    lnghWnd = WindowFromPoint(MousePoint.X, MousePoint.Y)
    lngDC = GetDC(lnghWnd)
    Call ScreenToClient(lnghWnd, MousePoint)
        
    lngResult = GetPixel(lngDC, MousePoint.X, MousePoint.Y)
    
    If lngResult <> -1 Then
        Colors = DefineRGB(lngResult)
        hscrColor(0).Value = Colors.Red
        hscrColor(1).Value = Colors.Green
        hscrColor(2).Value = Colors.Blue
    End If
End Sub

Private Sub PicColorPicker_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = 0
    Call ReleaseCapture
    bolSkipExample = False
    If Button = 1 Then
        PaintPointer PicColorPicker.BackColor, intSelPointer, True
        PaintExamples
    End If
End Sub

Private Sub picCustom_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Data.GetData(1) = 99 Then
        AddToCustom PointColors(), GradientType
    End If
End Sub

Private Sub picExample_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift = 1 Then
        If GradientType = Radial Then
            PaintGradient picExample, PointColors, GradientType, , X, Y
        Else
            lnAngle.X1 = X
            lnAngle.Y1 = Y
        End If
    End If
            
End Sub

Private Sub picExample_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Angle As Double
    If Button = 1 Then
        If GradientType = Linear And Shift = 1 Then
            lnAngle.X2 = X
            lnAngle.Y2 = Y
            lnAngle.Visible = True
            If lnAngle.Y1 > lnAngle.Y2 Then
                Angle = 0 - Atn((lnAngle.X1 - lnAngle.X2) / (lnAngle.Y1 - lnAngle.Y2)) / Trans
                If Angle < 0 Then
                    Angle = 360 + Angle
                End If
            ElseIf lnAngle.Y1 < lnAngle.Y2 Then
                Angle = 180 + Atn((lnAngle.X1 - lnAngle.X2) / (lnAngle.Y2 - lnAngle.Y1)) / Trans
            Else
                Angle = 90
            End If
            
'            If lnAngle.X1 < lnAngle.X2 Then
'                If lnAngle.Y2 <= lnAngle.Y1 Then
'                    Angle = 270 - Atn((lnAngle.Y2 - lnAngle.Y1) / (lnAngle.X1 - lnAngle.X2)) / Trans
'                Else
'                    Angle = 0 ' Atn((lnAngle.Y1 - lnAngle.Y2) / (lnAngle.X1 - lnAngle.X2)) / Trans
'                End If
'            End If
            txtAngle.Text = Round(Angle)
        ElseIf Me.Image <> 0 Then
            picExample.OLEDrag
        End If
    End If
End Sub

Private Sub picExample_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And GradientType = Linear Then PaintExamples
    lnAngle.Visible = False
End Sub

Private Sub picExample_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    Data.SetData 99
    AllowedEffects = vbDropEffectCopy
End Sub

Public Sub picIcon_Click(Index As Integer)
    bolCustomFocus = True
    picIconHolder.Left = picIcon(Index).Left - 36
    picIconHolder.Top = picIcon(Index).Top - 36
    picIconHolder.Visible = True
    intSelCustom = Index
    loadPointColors Index
    PaintExamples
End Sub

Private Sub picIcon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        picIconHolder.Left = picIcon(Index).Left - 36
        picIconHolder.Top = picIcon(Index).Top - 36
        picIconHolder.Visible = True
        bolCustomFocus = True
        intSelCustom = Index
        PopupMenu mnuPointer
    End If
End Sub

Private Sub picIcon_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Data.GetData(1) = 99 Then
        AddToCustom PointColors(), GradientType
    End If
End Sub

Private Sub picIconHolder_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Data.GetData(1) = 99 Then
        AddToCustom PointColors(), GradientType
    End If
End Sub

Private Sub picPointer_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim MousePoint As POINTAPI
Dim intOrig As Integer
    If Button = 1 Then
        'get mouse location
        Call GetCursorPos(MousePoint)
        MousePoint.X = MousePoint.X - (X - 5)
        'Set mouse location
        Call SetCursorPos(MousePoint.X, MousePoint.Y)
        bolLockPointer = True
        Call SetCapture(picPointer(Index).hWnd)
        
        If Index <> intSelPointer Then
            PaintPointer PointColors(picPointer(Index).Tag).Color, Index, True
            intOrig = intSelPointer
            intSelPointer = Index
            DoEvents
            If intOrig <> 0 Then
                PaintPointer PointColors(intOrig).Color, intOrig, False
            End If
        End If
        bolSkipExample = True
        hscrColor(0).Value = DefineRGB(PointColors(picPointer(Index).Tag).Color).Red
        hscrColor(1).Value = DefineRGB(PointColors(picPointer(Index).Tag).Color).Green
        bolSkipExample = True
        hscrColor(2).Value = DefineRGB(PointColors(picPointer(Index).Tag).Color).Blue
        bolCustomFocus = False
        picPointer(Index).ZOrder 0
    ElseIf Button = 2 Then
        If Index <> picPointer.UBound Then
            intSelPointer = Index
            PopupMenu mnuPointer
        End If
    End If
    
End Sub

Private Sub picPointer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim MousePoint As POINTAPI
Dim WindowPlace As RECT
Dim New_X As Long
Dim intCount As Integer

    If Button = 1 And bolLockPointer = False Then
        'get mouse location
        Call GetCursorPos(MousePoint)
        'get windowlocation
        Call GetWindowRect(Me.hWnd, WindowPlace)
        
        New_X = ((MousePoint.X - WindowPlace.Left - 2) * Screen.TwipsPerPixelX) - fraPointer.Left
        
        If New_X > 360 And New_X < 4360 Then
            If tbrSnap.Buttons("Snap").Value = tbrPressed Then
                For intCount = LBound(SnapPoints) To UBound(SnapPoints)
                    If New_X - 48 < SnapPoints(intCount) And New_X + 48 > SnapPoints(intCount) Then
                        'Set mouse location
                        Call SetCursorPos((SnapPoints(intCount) + fraPointer.Left + 38) / Screen.TwipsPerPixelX + WindowPlace.Left, MousePoint.Y)
                        New_X = SnapPoints(intCount)
                    End If
                Next
            End If
            picPointer(Index).Left = New_X - picPointer(Index).Width / 2
            PointColors(picPointer(Index).Tag).Place = ((New_X - picPointer(Index).Width / 2) - 360) / 40
            If picPointer(Index).Tag = picPointer.UBound Then
                NewPointer 0
            End If
        ElseIf New_X < 360 Then
            picPointer(Index).Left = 360 - picPointer(Index).Width / 2
            PointColors(picPointer(Index).Tag).Place = 0
        ElseIf New_X > 4360 Then
            picPointer(Index).Left = 4360 - picPointer(Index).Width / 2
            PointColors(picPointer(Index).Tag).Place = 100
        End If
        picIconHolder.Visible = False
        Me.Refresh
    Else
        bolLockPointer = False
    End If
    
End Sub

Private Sub DelPointer(Index As Integer)
Dim intCount As Integer

    picPointer(Index).Picture = picPointer(picPointer.UBound).Image
    picPointer(Index).Left = picPointer(picPointer.UBound).Left
    For intCount = Index To picPointer.UBound - 1
        PointColors(intCount) = PointColors(intCount + 1)
    Next
    For intCount = Index + 1 To picPointer.UBound
        picPointer(intCount).Tag = picPointer(intCount).Tag - 1
    Next
    Unload picPointer(picPointer.UBound)
    ReDim Preserve PointColors(1 To picPointer.UBound)
    intSelPointer = picPointer.UBound
    PaintExamples
End Sub

Public Sub NewPointer(Color As Long, Optional Place As Single = -1, Optional Selected As Boolean, Optional DontRedimPoints As Boolean)
    Load picPointer(picPointer.UBound + 1)
    If DontRedimPoints = False Then
        ReDim Preserve PointColors(1 To picPointer.UBound)
        PointColors(UBound(PointColors)).Place = Place
    End If
    picPointer(picPointer.UBound).Tag = picPointer.UBound
    PaintPointer Color, picPointer.UBound, Selected, DontRedimPoints
    picPointer(picPointer.UBound).Visible = True
    If Place <> -1 Then
        picPointer(picPointer.UBound).Left = 40 * Place + (360 - picPointer(picPointer.UBound).Width / 2)
    End If
End Sub

Private Sub picPointer_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    PaintExamples
End Sub

Private Sub txtDeg_GotFocus()
    txtAngle.SetFocus
End Sub

Private Sub txtAngle_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then PaintExamples
End Sub

Private Sub txtColorDec_GotFocus(Index As Integer)
    txtColorDec(Index).SelStart = 0
    txtColorDec(Index).SelLength = Len(txtColorDec(Index).Text)
End Sub

Private Sub txtColorDec_KeyPress(Index As Integer, KeyAscii As Integer)
'alleen getallen
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtColorDec_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Val(txtColorDec(Index).Text) < 0 Or txtColorDec(Index).Text = "" Then
        txtColorDec(Index).Text = 0
    ElseIf Val(txtColorDec(Index).Text) > 255 Then
        txtColorDec(Index).Text = 255
    End If
    hscrColor(Index).Value = txtColorDec(Index).Text
End Sub

Private Sub txtColorHex_GotFocus(Index As Integer)
    txtColorHex(Index).SelStart = 0
    txtColorHex(Index).SelLength = Len(txtColorDec(Index).Text)
End Sub

Private Sub txtColorHex_KeyPress(Index As Integer, KeyAscii As Integer)
'alleen getallen
    If KeyAscii >= 97 And KeyAscii <= 102 Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Exit Sub
    End If
    
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtColorHex_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim lngResult As Long

    If txtColorHex(Index).Text <> "" Then
        'The "&H" is placed in front of the value becouse it's a hex value
        lngResult = ("&H" & txtColorHex(Index).Text)
        
        If lngResult > 255 Then
            txtColorHex(Index).Text = "FF"
            lngResult = 255
        End If
        
        hscrColor(Index).Value = lngResult
        
    End If
End Sub
Public Sub PaintExamples()
        PaintGradient picGradient, PointColors, Linear
        PaintGradient picExample, PointColors, GradientType, Val(txtAngle.Text)
End Sub

Public Sub DrawSnapPointers(Number As Integer)
Dim intCount As Integer

    For intCount = 0 To Number
        lnSnapPointers(intCount).Visible = True
    Next
    
    For intCount = Number + 1 To 25
        lnSnapPointers(intCount).Visible = False
    Next
    
    ReDim SnapPoints(0 To Number)
    
    For intCount = 0 To Number
        SnapPoints(intCount) = 360 + ((4000 / Number) * intCount)
        lnSnapPointers(intCount).X1 = SnapPoints(intCount)
        lnSnapPointers(intCount).X2 = SnapPoints(intCount)
    Next
    
End Sub

Private Sub vscrCustom_Change()
    picCustom.Top = vscrCustom.Top + vscrCustom.Value
End Sub

Public Sub loadPointColors(Index)
Dim intCount As Integer
    
    PointColors() = CustomGradient(Index).Points()
    For intCount = picPointer.UBound To picPointer.LBound + 1 Step -1
        Unload picPointer(intCount)
    Next
    For intCount = LBound(PointColors) To UBound(PointColors) - 1
        NewPointer PointColors(intCount).Color, PointColors(intCount).Place, IIf(intCount = LBound(PointColors), True, False), True
    Next
    optLinear = IIf(CustomGradient(Index).FillType = Linear, True, False)
    optRadial = IIf(CustomGradient(Index).FillType = Radial, True, False)
    NewPointer 0
    
    intSelPointer = 1
    
    bolSkipExample = True
    hscrColor(0).Value = DefineRGB(PointColors(LBound(PointColors)).Color).Red
    hscrColor(1).Value = DefineRGB(PointColors(LBound(PointColors)).Color).Green
    bolSkipExample = True
    hscrColor(2).Value = DefineRGB(PointColors(LBound(PointColors)).Color).Blue
        
End Sub

Public Sub DelCustom(Index As Integer)
Dim intCount As Integer

    For intCount = Index To picIcon.UBound - 1
        picIcon(intCount).Picture = picIcon(intCount + 1).Image
    Next
    
    Unload picIcon(picIcon.UBound)
    
    For intCount = Index To UBound(CustomGradient) - 1
        CustomGradient(intCount) = CustomGradient(intCount + 1)
    Next
    
    ReDim Preserve CustomGradient(1 To UBound(CustomGradient) - 1)
    
    If bolCustomFocus = True And UBound(CustomGradient) + 1 <> intSelCustom Then
        picIconHolder.Left = picIcon(intSelCustom).Left - 36
        picIconHolder.Top = picIcon(intSelCustom).Top - 36
        picIconHolder.Visible = True
    Else
        picIconHolder.Visible = False
    End If
    
    If picIcon(UBound(CustomGradient)).Top + _
     picIcon(UBound(CustomGradient)).Height + 120 < picCustom.Height Then
        picCustom.Height = picIcon(UBound(CustomGradient)).Top + _
        picIcon(UBound(CustomGradient)).Height + 210
        If picCustom.Height <> 1110 Then
            vscrCustom.Min = 0
            vscrCustom.Max = vscrCustom.Height - picCustom.Height
            vscrCustom.Value = vscrCustom.Max
            vscrCustom.Enabled = True
        Else
            picCustom.Top = 100
            vscrCustom.Enabled = False
        End If
    End If
    
    
    
    Me.Refresh
End Sub

Public Sub NewCustom()
Dim intCount As Integer

    ReDim CustomGradient(0 To 0)
    On Error Resume Next
    For intCount = 1 To picIcon.UBound
        Unload picIcon(intCount)
    Next
    picIconHolder.Visible = False
End Sub

Public Function Asn(Number As Double) As Double
    Asn = Atn(Number / Sqr(-Number * Number + 1))
End Function
