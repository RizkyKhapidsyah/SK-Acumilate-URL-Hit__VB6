VERSION 5.00
Begin VB.Form FrmAdd 
   Appearance      =   0  'Flat
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add URL"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7245
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   30
      ScaleHeight     =   945
      ScaleWidth      =   7155
      TabIndex        =   0
      Top             =   30
      Width           =   7185
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   90
         TabIndex        =   1
         Text            =   "http://www.cbel.com/Lottery_Gambling?p=6730&s=13&l=13"
         Top             =   510
         Width           =   5685
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter or paste a web link and click Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   120
         TabIndex        =   4
         Top             =   150
         Width           =   4680
      End
      Begin VB.Label CommandBack 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Close"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6000
         MouseIcon       =   "FrmAdd.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   120
         Width           =   945
      End
      Begin VB.Label Command9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Add"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6000
         MouseIcon       =   "FrmAdd.frx":0152
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   510
         Width           =   945
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         Height          =   315
         Index           =   0
         Left            =   5970
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   1005
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         Height          =   315
         Index           =   1
         Left            =   5970
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Width           =   1005
      End
   End
End
Attribute VB_Name = "FrmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command9_Click()
For p = 0 To FrmMain.List1.ListCount - 1
If FrmMain.List1.List(p) = Text5.Text Then
MsgBox "Site already exists"
Exit Sub
End If
Next p
FrmMain.List1.AddItem Text5.Text
FrmAdd.Visible = False
FrmAdd.Cls
End Sub

Private Sub CommandBack_Click()
FrmAdd.Visible = False
FrmAdd.Cls
End Sub

Private Sub Form_Load()
SetWindowPos FrmAdd.hwnd, conHwndTopmost, 100, 100, 490, 98, conSwpNoActivate Or conSwpShowWindow

End Sub
