VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmHTML 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ASCII Art Generator"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   406
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   538
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtRepeat 
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   36
      Top             =   5160
      Width           =   7455
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   1815
      TabIndex        =   33
      Top             =   4680
      Width           =   1815
      Begin VB.OptionButton OptText 
         Caption         =   "Repeated"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   35
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton OptText 
         Caption         =   "Random"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CheckBox ChkBlue 
      Caption         =   "Use Blue"
      Height          =   255
      Left            =   4920
      TabIndex        =   31
      Top             =   3480
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox ChkGreen 
      Caption         =   "Use Green"
      Height          =   255
      Left            =   4920
      TabIndex        =   30
      Top             =   3240
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox ChkRed 
      Caption         =   "Use Red"
      Height          =   255
      Left            =   4920
      TabIndex        =   29
      Top             =   3000
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.TextBox TxtTitle 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   3840
      Width           =   7815
   End
   Begin VB.CommandButton CmdPreview 
      Caption         =   "View Preview"
      Height          =   375
      Left            =   1920
      TabIndex        =   26
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton CmdBackColour 
      Caption         =   "Change..."
      Height          =   375
      Left            =   6600
      TabIndex        =   25
      Top             =   1800
      Width           =   1335
   End
   Begin VB.PictureBox PicColour 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   6120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   24
      Top             =   1800
      Width           =   375
   End
   Begin VB.ComboBox CmbFont 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   3240
      Width           =   1695
   End
   Begin VB.PictureBox PicLoad 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   8640
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   17
      Top             =   360
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   7560
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   4560
      ScaleHeight     =   1215
      ScaleWidth      =   1455
      TabIndex        =   14
      Top             =   1800
      Width           =   1455
      Begin VB.OptionButton OptColour 
         Caption         =   "Greyscale"
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   22
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton OptColour 
         Caption         =   "Blue"
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   19
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton OptColour 
         Caption         =   "Green"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   18
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton OptColour 
         Caption         =   "Red"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton OptColour 
         Caption         =   "Original Colours"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton CmdGo 
      Caption         =   "Create"
      Height          =   375
      Left            =   6600
      TabIndex        =   12
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox TxtSize 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Text            =   "8"
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox TxtY 
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Text            =   "60"
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox TxtX 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Text            =   "150"
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton CmdHTMLBrowse 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   6600
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox TxtHTML 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1080
      Width           =   6375
   End
   Begin VB.CommandButton CmdPicBrowse 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   6600
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox TxtPic 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   6375
   End
   Begin VB.Label Label10 
      Caption         =   "Text:"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Title:"
      Height          =   195
      Left            =   120
      TabIndex        =   28
      Top             =   3600
      Width           =   345
   End
   Begin VB.Label Label8 
      Caption         =   "Background Colour:"
      Height          =   255
      Left            =   6120
      TabIndex        =   23
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Font:"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "Colour Style:"
      Height          =   255
      Left            =   4560
      TabIndex        =   13
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Font Size:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "X"
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "Output Res:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Output HTML:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Picture File:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   825
   End
End
Attribute VB_Name = "FrmHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal IntX As Long, ByVal IntY As Long) As Long

Dim HTMLOutput As String

Private Sub CmdBackColour_Click()
    CD.ShowColor
    PicColour.BackColor = CD.Color
End Sub

Private Sub CmdGo_Click()
    Dim RanTxt As Boolean
    If TxtPic.Text = "" Then
        Call MsgBox("You require a picture File", vbInformation)
        Exit Sub
    End If
    If TxtHTML.Text = "" Then
        Call MsgBox("You require an output File", vbInformation)
        Exit Sub
    End If
    If TxtX.Text = "" Or Val(TxtX.Text) = 0 Then
        Call MsgBox("Your Width is invalid", vbInformation)
        Exit Sub
    End If
    If TxtY.Text = "" Or Val(TxtY.Text) = 0 Then
        Call MsgBox("Your Height is invalid", vbInformation)
        Exit Sub
    End If
    If TxtSize.Text = "" Or Val(TxtSize.Text) = 0 Then
        Call MsgBox("Your fontsize is invalid", vbInformation)
        Exit Sub
    End If
    RanTxt = True
    If OptText(1).Value = True Then
        RanTxt = False
        If TxtRepeat.Text = "" Then
            Call MsgBox("Text is required to repeat", vbInformation)
            Exit Sub
        End If
    End If
    Dim RE As Boolean
    Dim GE As Boolean
    Dim BE As Boolean
    Dim Cols As Integer
    Dim Render As Integer
    
    Cols = 0
    If ChkRed.Value = 1 Then
        RE = True
        Cols = Cols + 1
    End If
    If ChkGreen.Value = 1 Then
        GE = True
        Cols = Cols + 1
    End If
    If ChkBlue.Value = 1 Then
        BE = True
        Cols = Cols + 1
    End If
    If Cols = 0 Then
        Call MsgBox("Please select which colour sources to use", vbInformation)
        Exit Sub
    End If
    
    Dim Colour As Long
    Dim Red As Integer
    Dim Green As Integer
    Dim Blue As Integer
    Dim Av As Long
    
    If OptColour(0).Value = True Then
        Render = 0
    ElseIf OptColour(1).Value = True Then
        Render = 1
    ElseIf OptColour(2).Value = True Then
        Render = 2
    ElseIf OptColour(3).Value = True Then
        Render = 3
    ElseIf OptColour(4).Value = True Then
        Render = 4
    End If
    
    
    HTMLOutput = ""
    WebCol = DectoWebCol(PicColour.BackColor)
    HTMLOutput = "<HTML><HEAD><Title>Kevin Pfisters ASCII ART - " & TxtTitle.Text & "</title><CENTER><FONT face=" & """" & CmbFont & """" & " size=" & TxtSize.Text & "><Body bgcolor=" & """" & WebCol & """" & ">"
    PicLoad.Picture = LoadPicture(TxtPic.Text)
    Open TxtHTML.Text For Output As #1
    Print #1, HTMLOutput
    HTMLOutput = ""
    Dim Z As Long
    For Y = 1 To Val(TxtY.Text)
        For X = 1 To Val(TxtX.Text)
            Z = Z + 1
            Colour = GetPixel(PicLoad.hdc, (PicLoad.Width - 1) / Val(TxtX.Text) * X, (PicLoad.Height - 1) / Val(TxtY.Text) * Y)
            If Render > 0 Then
                GetRgb Colour, Red, Green, Blue
                Av = 0
                If RE = True Then
                    Av = Av + Red
                End If
                If GE = True Then
                    Av = Av + Green
                End If
                If BE = True Then
                    Av = Av + Blue
                End If
                Av = Av / Cols
            End If
            If Render = 0 Then
                WebCol = DectoWebCol(Colour)
            ElseIf Render = 1 Then
                WebCol = DectoWebCol(RGB(Av, 0, 0))
            ElseIf Render = 2 Then
                WebCol = DectoWebCol(RGB(0, Av, 0))
            ElseIf Render = 3 Then
                WebCol = DectoWebCol(RGB(0, 0, Av))
            ElseIf Render = 4 Then
                WebCol = DectoWebCol(RGB(Av, Av, Av))
            End If
            HTMLOutput = HTMLOutput & "<font color=" & """" & WebCol & """" & "font STYLE=" & """font-size: " & Val(TxtSize.Text) & "px" & """" & ">"
            If RanTxt = True Then
                HTMLOutput = HTMLOutput & Chr(Rnd * 26 + 64) & "</font>"
            Else
                HTMLOutput = HTMLOutput & Mid(TxtRepeat.Text, (Z Mod Len(TxtRepeat.Text) + 1), 1) & "</font>"
            End If
        Next
        HTMLOutput = HTMLOutput & "<br>" & vbNewLine
        Print #1, HTMLOutput
        HTMLOutput = ""
        Me.Caption = "ASCII Art Generator ~ " & Int(100 / Val(TxtY.Text) * Y) & "%"
        DoEvents
    Next
    HTMLOutput = HTMLOutput & "<br></FONT></CENTER></BODY></HTML>"
    Print #1, HTMLOutput
    Close
    Me.Caption = "ASCII Art Generator"
    ask = MsgBox("Completed Conversion" & vbNewLine & "Would you like to see the output HTML?", vbYesNo)
    If ask = vbYes Then
        Dim BrowserString As String
        BrowserString = "rundll32.exe url.dll,FileProtocolHandler " & TxtHTML.Text
        Search = Shell(BrowserString, 0)
    End If
End Sub

Private Sub CmdHTMLBrowse_Click()
    CD.Filter = "HTML (*.HTM)|*.HTM|Any File (*.*)|*.*"
    CD.ShowSave
    FileName = CD.FileName
    If FileName = "" Then Exit Sub
    TxtHTML.Text = FileName
End Sub

Private Sub CmdPicBrowse_Click()
    CD.Filter = "Bitmap (*.BMP)|*.BMP|JPEG (*.JPG)|*.JPG|GIF (*.GIF)|*.GIF|Any File (*.*)|*.*"
    CD.ShowOpen
    FileName = CD.FileName
    If FileName = "" Then Exit Sub
    TxtPic.Text = FileName
End Sub

Sub GetRgb(ByVal Color As Long, ByRef Red As Integer, ByRef Green As Integer, ByRef Blue As Integer)
    Dim LngColVal As Long
    LngColVal = Color And 255
    Red = LngColVal And 255
    LngColVal = Int(Color / 256)
    Green = LngColVal And 255
    LngColVal = Int(Color / 65536)
    Blue = LngColVal And 255
End Sub

Public Function DectoWebCol(lngColour As Long) As String
    Dim strColour As String
    strColour = Hex(lngColour)
    Do While Len(strColour) < 6
        strColour = "0" & strColour
    Loop
    DectoWebCol = "#" & Right$(strColour, 2) & _
    Mid$(strColour, 3, 2) & _
    Left$(strColour, 2)
End Function

Private Sub CmdPreview_Click()
    Me.Caption = "ASCII Art Generator ~ Generating Preview"
    HTMLOutput = ""
    WebCol = DectoWebCol(PicColour.BackColor)
    HTMLOutput = "<HTML><HEAD><Title>Preview</title><CENTER><FONT face=" & """" & CmbFont & """" & " size=" & TxtSize.Text & "><Body bgcolor=" & """" & WebCol & """" & ">"
    Open "C:\TmpHTML.htm" For Output As #1
    Print #1, HTMLOutput
    HTMLOutput = ""
    Dim Av As Long
    For Y = 1 To Val(TxtY.Text)
        For X = 1 To Val(TxtX.Text)
            Av = 128 / Val(TxtX.Text) * X + 128 / Val(TxtY.Text) * Y
            WebCol = DectoWebCol(RGB(Av, Av, Av))
            HTMLOutput = HTMLOutput & "<font color=" & """" & WebCol & """" & "font STYLE=" & """font-size: " & Val(TxtSize.Text) & "px" & """" & ">" & Chr(Rnd * 26 + 64) & "</font>"
        Next
        HTMLOutput = HTMLOutput & "<br>" & vbNewLine
        Print #1, HTMLOutput
        HTMLOutput = ""
        DoEvents
    Next
    HTMLOutput = HTMLOutput & "<br></FONT></CENTER></BODY></HTML>"
    Print #1, HTMLOutput
    Close
    
    B = Timer
    Do
        DoEvents
    Loop Until Timer - B > 1
    
    BrowserString = "rundll32.exe url.dll,FileProtocolHandler " & "C:\TmpHTML.htm"
    Search = Shell(BrowserString, 0)
    DoEvents
    B = Timer
    Do
        DoEvents
    Loop Until Timer - B > 1
    
    Kill "C:\TmpHTML.htm"
    Me.Caption = "ASCII Art Generator"
End Sub

Private Sub Form_Load()
    For X = 0 To Screen.FontCount - 1
        CmbFont.AddItem Screen.Fonts(X)
    Next
    CmbFont.Text = "Terminal"
End Sub

Private Sub OptText_Click(Index As Integer)
    If Index = 1 Then
        TxtRepeat.Enabled = True
    Else
        TxtRepeat.Enabled = False
    End If
End Sub
