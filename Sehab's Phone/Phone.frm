VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Caller"
   ClientHeight    =   6435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Phone.frx":0000
   ScaleHeight     =   6435
   ScaleWidth      =   2505
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList NumberList 
      Left            =   1080
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   30
      ImageHeight     =   29
      MaskColor       =   128
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Phone.frx":34CDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Phone.frx":35798
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Phone.frx":36256
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Phone.frx":36D14
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Phone.frx":377D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Phone.frx":38290
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Phone.frx":38D4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Phone.frx":3980C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Phone.frx":3A2CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Phone.frx":3AD88
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Phone.frx":3B846
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Phone.frx":3C304
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Phone.frx":3CDC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Phone.frx":3D880
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Phone.frx":3E33E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Phone.frx":3EDFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Phone.frx":3F8BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Phone.frx":40378
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Phone.frx":40E36
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Phone.frx":418F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Phone.frx":423B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Phone.frx":42E70
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Phone.frx":4392E
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Phone.frx":443EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Number 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   200
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Call"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Help 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   6120
      Width           =   255
   End
   Begin VB.Image Delete 
      Height          =   195
      Left            =   1800
      Picture         =   "Phone.frx":44EAA
      Top             =   1320
      Width           =   450
   End
   Begin VB.Shape Borderer 
      BorderColor     =   &H00FF0000&
      Height          =   255
      Left            =   2040
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label M 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.Label XX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Num12 
      Height          =   435
      Left            =   1680
      Picture         =   "Phone.frx":45398
      Top             =   3960
      Width           =   450
   End
   Begin VB.Image Num11 
      Height          =   435
      Left            =   1030
      Picture         =   "Phone.frx":45E46
      Top             =   3960
      Width           =   450
   End
   Begin VB.Image Num10 
      Height          =   435
      Left            =   370
      Picture         =   "Phone.frx":468F4
      Top             =   3960
      Width           =   450
   End
   Begin VB.Image Num9 
      Height          =   435
      Left            =   1680
      Picture         =   "Phone.frx":473A2
      Top             =   3360
      Width           =   450
   End
   Begin VB.Image Num8 
      Height          =   435
      Left            =   1030
      Picture         =   "Phone.frx":47E50
      Top             =   3370
      Width           =   450
   End
   Begin VB.Image Num7 
      Height          =   435
      Left            =   360
      Picture         =   "Phone.frx":488FE
      Top             =   3360
      Width           =   450
   End
   Begin VB.Image Num6 
      Height          =   435
      Left            =   1680
      Picture         =   "Phone.frx":493AC
      Top             =   2775
      Width           =   450
   End
   Begin VB.Image Num5 
      Height          =   435
      Left            =   1030
      Picture         =   "Phone.frx":49E5A
      Top             =   2790
      Width           =   450
   End
   Begin VB.Image Num4 
      Height          =   435
      Left            =   360
      Picture         =   "Phone.frx":4A908
      Top             =   2760
      Width           =   450
   End
   Begin VB.Image Num3 
      Height          =   435
      Left            =   1680
      Picture         =   "Phone.frx":4B3B6
      Top             =   2190
      Width           =   450
   End
   Begin VB.Image Num2 
      Height          =   435
      Left            =   1030
      Picture         =   "Phone.frx":4BE64
      Top             =   2190
      Width           =   450
   End
   Begin VB.Image Num1 
      Height          =   435
      Left            =   360
      Picture         =   "Phone.frx":4C912
      Top             =   2190
      Width           =   450
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Teller As New SpVoice
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function tapiRequestMakeCall Lib "TAPI32.DLL" _
    (ByVal DestAddr$, ByVal AppName As String, _
    ByVal CalledParty As String, ByVal Comment As String) As Long

Dim strNumber As String
Dim PosX As Integer
Dim PosY As Integer
Dim strNew As String
    
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Sub Command1_Click()
    
    If strNew = "" Then
        Teller.Speak "No Number To Dial": Exit Sub
    Else
        Teller.Speak "Dialing: " & strNew
        tapiRequestMakeCall Number.Text, App.Title, "Dialed", "Hey"
    End If
    
End Sub

Private Sub Delete_Click()
    If Number.Text = "" Then Beep: Exit Sub
    strNumber = Mid(Number.Text, 1, Len(Number.Text) - 1)
    strNew = Mid(strNew, 1, Len(strNew) - 2)
    Number.Text = strNumber
End Sub

Private Sub Form_Load()
    Num1.Picture = NumberList.ListImages.Item(1).ExtractIcon
    Num2.Picture = NumberList.ListImages.Item(3).ExtractIcon
    Num3.Picture = NumberList.ListImages.Item(5).ExtractIcon
    Num4.Picture = NumberList.ListImages.Item(7).ExtractIcon
    Num5.Picture = NumberList.ListImages.Item(9).ExtractIcon
    Num6.Picture = NumberList.ListImages.Item(11).ExtractIcon
    Num7.Picture = NumberList.ListImages.Item(13).ExtractIcon
    Num8.Picture = NumberList.ListImages.Item(15).ExtractIcon
    Num9.Picture = NumberList.ListImages.Item(17).ExtractIcon
    Num10.Picture = NumberList.ListImages.Item(21).ExtractIcon
    Num11.Picture = NumberList.ListImages.Item(19).ExtractIcon
    Num12.Picture = NumberList.ListImages.Item(23).ExtractIcon
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PosX = X
    PosY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Borderer.Visible = False
    M.ForeColor = vbYellow
    XX.ForeColor = vbYellow
    Help.ForeColor = vbYellow
    
    Dim Z As POINTAPI
    GetCursorPos Z
    
    If Button = 1 Then
        Me.Left = (Z.X * 15) - PosX
        Me.Top = (Z.Y * 15) - PosY
    End If
        
End Sub

Private Sub Help_Click()
    Teller.Speak "Ok so you need help, eh?  All you do is press in the numbers of the telephone number, and then press call.  But you can't be on the internet.  That's Easy Isn't It?"
    MsgBox "Is That Easy?", vbYesNo
End Sub

Private Sub Help_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Borderer.Visible = True
    Borderer.Left = Help.Left
    Borderer.Top = Help.Top
    Help.ForeColor = vbRed
    M.ForeColor = vbYellow
    XX.ForeColor = vbYellow
End Sub

Private Sub M_Click()
    Me.WindowState = vbMinimized
    M.ForeColor = vbYellow
    Borderer.Visible = False
End Sub

Private Sub M_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Borderer.Visible = True
    Borderer.Left = M.Left
    Borderer.Top = M.Top
    M.ForeColor = vbRed
    XX.ForeColor = vbYellow
    Help.ForeColor = vbYellow
End Sub

'number 1
Private Sub Num1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Num1.Picture = NumberList.ListImages.Item(2).ExtractIcon
    Number.Text = Number.Text & "1"
    strNew = strNew & "1 "
End Sub

Private Sub Num1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Num1.Picture = NumberList.ListImages.Item(1).ExtractIcon
End Sub

'number 2
Private Sub Num2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Num2.Picture = NumberList.ListImages.Item(4).ExtractIcon
    Number.Text = Number.Text & "2"
    strNew = strNew & "2 "
End Sub

Private Sub Num2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Num2.Picture = NumberList.ListImages.Item(3).ExtractIcon
End Sub

'number 3
Private Sub Num3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Num3.Picture = NumberList.ListImages.Item(6).ExtractIcon
    Number.Text = Number.Text & "3"
    strNew = strNew & "3 "
End Sub

Private Sub Num3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Num3.Picture = NumberList.ListImages.Item(5).ExtractIcon
End Sub

'number 4
Private Sub Num4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Num4.Picture = NumberList.ListImages.Item(8).ExtractIcon
    Number.Text = Number.Text & "4"
    strNew = strNew & "4 "
End Sub

Private Sub Num4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Num4.Picture = NumberList.ListImages.Item(7).ExtractIcon
End Sub

'number 5
Private Sub Num5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Num5.Picture = NumberList.ListImages.Item(10).ExtractIcon
    Number.Text = Number.Text & "5"
    strNew = strNew & "5 "
End Sub

Private Sub Num5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Num5.Picture = NumberList.ListImages.Item(9).ExtractIcon
End Sub

'number 6
Private Sub Num6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Num6.Picture = NumberList.ListImages.Item(12).ExtractIcon
    Number.Text = Number.Text & "6"
    strNew = strNew & "6 "
End Sub

Private Sub Num6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Num6.Picture = NumberList.ListImages.Item(11).ExtractIcon
End Sub

'number 7
Private Sub Num7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Num7.Picture = NumberList.ListImages.Item(14).ExtractIcon
    Number.Text = Number.Text & "7"
    strNew = strNew & "7 "
End Sub

Private Sub Num7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Num7.Picture = NumberList.ListImages.Item(13).ExtractIcon
End Sub

'number 8
Private Sub Num8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Num8.Picture = NumberList.ListImages.Item(16).ExtractIcon
    Number.Text = Number.Text & "8"
    strNew = strNew & "8 "
End Sub

Private Sub Num8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Num8.Picture = NumberList.ListImages.Item(15).ExtractIcon
End Sub

'number 9
Private Sub Num9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Num9.Picture = NumberList.ListImages.Item(18).ExtractIcon
    Number.Text = Number.Text & "9"
    strNew = strNew & "9 "
End Sub

Private Sub Num9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Num9.Picture = NumberList.ListImages.Item(17).ExtractIcon
End Sub

'number 10
Private Sub Num10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Num10.Picture = NumberList.ListImages.Item(22).ExtractIcon
    Number.Text = Number.Text & "*"
    strNew = strNew & "* "
End Sub

Private Sub Num10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Num10.Picture = NumberList.ListImages.Item(21).ExtractIcon
End Sub

'number 11
Private Sub Num11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Num11.Picture = NumberList.ListImages.Item(20).ExtractIcon
    Number.Text = Number.Text & "0"
    strNew = strNew & "0 "
End Sub

Private Sub Num11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Num11.Picture = NumberList.ListImages.Item(19).ExtractIcon
End Sub

'number 12
Private Sub Num12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Num12.Picture = NumberList.ListImages.Item(24).ExtractIcon
    Number.Text = Number.Text & "#"
    strNew = strNew & "# "
End Sub

Private Sub Num12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Num12.Picture = NumberList.ListImages.Item(23).ExtractIcon
End Sub

Private Sub XX_Click()
    End
End Sub

Private Sub XX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Borderer.Visible = True
    Borderer.Left = XX.Left
    Borderer.Top = XX.Top
    XX.ForeColor = vbRed
    M.ForeColor = vbYellow
End Sub
