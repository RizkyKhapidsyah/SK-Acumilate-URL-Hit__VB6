VERSION 5.00
Begin VB.Form FrmCookies 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cookies"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2550
      Left            =   60
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   60
      Width           =   4965
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C000&
      Caption         =   "Keep password cookies"
      Height          =   225
      Left            =   5100
      TabIndex        =   1
      Top             =   1350
      Value           =   1  'Checked
      Width           =   2145
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0C000&
      Caption         =   "Delete all cookies"
      Height          =   255
      Left            =   5100
      TabIndex        =   0
      Top             =   1650
      Width           =   2145
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   $"FrmCookies.frx":0000
      Height          =   975
      Index           =   2
      Left            =   5130
      TabIndex        =   5
      Top             =   90
      Width           =   2955
   End
   Begin VB.Label Command5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Delete Cookies"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6750
      MouseIcon       =   "FrmCookies.frx":00A0
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2310
      Width           =   1305
   End
   Begin VB.Label Command4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Get Cookies"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5220
      MouseIcon       =   "FrmCookies.frx":01F2
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2310
      Width           =   1305
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   1
      Left            =   6720
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   1365
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   0
      Left            =   5190
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   1365
   End
End
Attribute VB_Name = "FrmCookies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Check2.Value = 1 Then Check2.Value = 0
End Sub

Private Sub Check2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Check1.Value = 1 Then Check1.Value = 0
End Sub

Private Sub Command4_Click()
FrmMain.GetCacheURLList

If List2.ListCount > 0 Then
Command5.Visible = True
Else
MsgBox "cookies empty"
End If

End Sub

Private Sub Command5_Click()

  Dim cachefile As String
   Dim I As Long
     
  'delete all files except..
   For I = 0 To FrmCookies.List2.ListCount - 1
   
      cachefile = FrmCookies.List2.List(I)
      
     '..if the file is a cookie, don't screw
     'up saved passwords, so skip it
     If Check1.Value = 1 Then
      If InStr(cachefile, "Cookie") = 0 Then
         Call DeleteUrlCacheEntry(cachefile)
      End If
     End If
     
     If Check2.Value = 1 Then
   For I2 = 0 To List2.ListCount - 1
   cachefile = List2.List(I2)
   Call DeleteUrlCacheEntry(cachefile)
   Next I2
   
  'reload the list
   FrmMain.GetCacheURLList
     End If
     
     
   Next
   
  'reload the list
   FrmMain.GetCacheURLList
Command4_Click
End Sub

Private Sub Form_Load()
SetWindowPos FrmCookies.hwnd, conHwndTopmost, 400, 400, 550, 210, conSwpNoActivate Or conSwpShowWindow

End Sub
