VERSION 5.00
Begin VB.Form FrmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Web site download / hit counter and vistor increaser"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   3960
      Top             =   3450
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Use all links listed above by turn"
      Height          =   225
      Left            =   60
      TabIndex        =   12
      Top             =   4560
      Width           =   2565
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3090
      ItemData        =   "hit.frx":0000
      Left            =   30
      List            =   "hit.frx":000A
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   30
      Width           =   5835
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   780
      TabIndex        =   6
      Text            =   "350"
      Top             =   3810
      Width           =   705
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Left            =   780
      TabIndex        =   5
      Text            =   "0"
      Top             =   3510
      Width           =   705
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   780
      TabIndex        =   3
      Text            =   "50"
      Top             =   3210
      Width           =   705
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Text            =   "http://www.cbel.com/Lottery_Gambling?p=6730&s=13&l=13"
      Top             =   4200
      Width           =   5235
   End
   Begin VB.Label Command2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stop"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5940
      MouseIcon       =   "hit.frx":007C
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   4560
      Width           =   945
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Start"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7020
      MouseIcon       =   "hit.frx":01CE
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   4560
      Width           =   945
   End
   Begin VB.Label Command4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cookies"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5970
      MouseIcon       =   "hit.frx":0320
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   2820
      Width           =   2025
   End
   Begin VB.Label Command7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Load url's"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6300
      MouseIcon       =   "hit.frx":0472
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   2070
      Width           =   1365
   End
   Begin VB.Label Command8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Save url's"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6300
      MouseIcon       =   "hit.frx":05C4
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   1710
      Width           =   1365
   End
   Begin VB.Label Command6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Delete"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7050
      MouseIcon       =   "hit.frx":0716
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   2460
      Width           =   945
   End
   Begin VB.Label Command3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Add"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5970
      MouseIcon       =   "hit.frx":0868
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   2460
      Width           =   945
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Current hit"
      Height          =   255
      Index           =   4
      Left            =   1560
      TabIndex        =   11
      Top             =   3510
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the amount of times you wish to hit a web link"
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   10
      Top             =   3210
      Width           =   5835
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Some sites will not register a hit if this is below 300 others above 1500"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   8
      Top             =   3810
      Width           =   5835
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Timer ="
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   3810
      Width           =   705
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Count"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3510
      Width           =   1245
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Stop at "
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3210
      Width           =   1245
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Web address to hit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   1
      Top             =   4170
      Width           =   4965
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   2
      Left            =   5940
      Shape           =   4  'Rounded Rectangle
      Top             =   2430
      Width           =   1005
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   3
      Left            =   7020
      Shape           =   4  'Rounded Rectangle
      Top             =   2430
      Width           =   1005
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   4
      Left            =   6270
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   1425
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   5
      Left            =   6270
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   1425
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   0
      Left            =   5940
      Shape           =   4  'Rounded Rectangle
      Top             =   2790
      Width           =   2085
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   1
      Left            =   5910
      Shape           =   4  'Rounded Rectangle
      Top             =   4530
      Width           =   1005
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   6
      Left            =   6990
      Shape           =   4  'Rounded Rectangle
      Top             =   4530
      Width           =   1005
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal numBytes As Long)




Private Sub Command1_Click()
Timer1.Interval = Val(Text4.Text)

If Check3.Value = 1 Then List1.Selected(0) = True

Timer1.Enabled = True

SaveSetting "Hit", "site", "as", Text1.Text
End Sub

Function Run(strFilePath As String, Optional strParms As String, Optional strDir As String) As String
       
  Const SW_SHOW = 5

  Select Case ShellExecute(0, "Open", strFilePath, strParms, strDir, SW_SHOW)
  Case 0
    Run = "Insufficent system memory or corrupt program file"
  Case Else
    Run = ""
  End Select

End Function


Private Sub Command2_Click()
Timer1.Enabled = False
Text3.Text = "0"
FrmMain.ZOrder 0
End Sub

Private Sub Command3_Click()
centa FrmAdd
FrmAdd.Visible = True
FrmAdd.SetFocus
End Sub


Public Sub GetCacheURLList()
    
   Dim ICEI As INTERNET_CACHE_ENTRY_INFO
   Dim hFile As Long
   Dim cachefile As String
   Dim posUrl As Long
   Dim posEnd As Long
   Dim dwBuffer As Long
   Dim pntrICE As Long
   
  FrmCookies.List2.Clear
   
  'Like other APIs, calling FindFirstUrlCacheEntry or
  'FindNextUrlCacheEntry with an insufficient buffer will
  'cause the API to fail, and the buffer pointing to the
  'correct size required for a successful call.
   dwBuffer = 0

  'Call to determine the required buffer size
   hFile = FindFirstUrlCacheEntry(0&, ByVal 0, dwBuffer)
   
  'both conditions hould be met by the first call
   If (hFile = ERROR_CACHE_FIND_FAIL) And _
      (Err.LastDllError = ERROR_INSUFFICIENT_BUFFER) Then
   
     'The INTERNET_CACHE_ENTRY_INFO data type is a
     'variable-length type. It is neccessary to allocate
     'memnory for the result of the call and pass the
     'pointer to this memory location to the API.
      pntrICE = LocalAlloc(LMEM_FIXED, dwBuffer)
        
     'allocation successful
      If pntrICE Then
         
        'set a Long pointer to the memory location
         CopyMemory ByVal pntrICE, dwBuffer, 4
         
        'and call the first find API again passing the
        'pointer to the allocated memory
         hFile = FindFirstUrlCacheEntry(vbNullString, ByVal pntrICE, dwBuffer)
       
        'hfile should = 1 (success)
         If hFile <> ERROR_CACHE_FIND_FAIL Then
         
           'loop through the cache
            Do
            
              'the pointer has ben filled, so move the
              'data back into a ICEI structure
               CopyMemory ICEI, ByVal pntrICE, Len(ICEI)
            
              'CacheEntryType is a long representing
              'the type of entry returned
               If (ICEI.CacheEntryType And _
                   NORMAL_CACHE_ENTRY) = NORMAL_CACHE_ENTRY Then
               
                 'extract the string from the memory location
                 'pointed to by the lpszSourceUrlName member
                 'and add to a list
                  cachefile = GetStrFromPtrA(ICEI.lpszSourceUrlName)
                  
                  
                  FrmCookies.List2.AddItem cachefile

               End If
               
              'free the pointer and memory associated
              'with the last-retrieved file
               Call LocalFree(pntrICE)
               
              'and again repeat the procedure, this time calling
              'FindNextUrlCacheEntry with a buffer size set to 0.
              'This will cause the call to once again fail,
              'returning the required size as dwBuffer
               dwBuffer = 0
               Call FindNextUrlCacheEntry(hFile, ByVal 0, dwBuffer)
               
              'allocate and assign the memory to the pointer
               pntrICE = LocalAlloc(LMEM_FIXED, dwBuffer)
               CopyMemory ByVal pntrICE, dwBuffer, 4
               
           'and call again with the valid parameters.
           'If the call fails (no more data), the loop exits.
           'If the call is successful, the Do portion of the
           'loop is executed again, extracting the data from
           'the returned type
            Loop While FindNextUrlCacheEntry(hFile, ByVal pntrICE, dwBuffer)
  
         End If 'hFile
         
      End If 'pntrICE
   
   End If 'hFile
   
  'clean up by closing the find handle, as
  'well as calling LocalFree again to be safe
   Call LocalFree(pntrICE)
   Call FindCloseUrlCache(hFile)
   
End Sub



Private Sub Command4_Click()
FrmCookies.Show
centa FrmCookies
End Sub

Private Sub Command6_Click()
If LenB(List1.Text) = 0 Then
MsgBox "Please select a web link to delete", vbOKOnly, App.Title
Exit Sub
End If
List1.RemoveItem List1.ListIndex
Call SaveListBox(App.Path & "\urls.txt", List1)
End Sub

Private Sub Command7_Click()
List1.Clear
Call Loadlistbox(App.Path & "\urls.txt", List1)
End Sub

Private Sub Command8_Click()
Call SaveListBox(App.Path & "\urls.txt", List1)
End Sub
Public Sub SaveListBox(Directory As String, TheList As ListBox)
    Dim savelist As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For savelist& = 0 To TheList.ListCount - 1
        Print #1, TheList.List(savelist&)
    Next savelist&
    Close #1
End Sub
Public Sub Loadlistbox(Directory As String, TheList As ListBox)
    Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    
    While Not EOF(1)
        Input #1, MyString$
        If LOF(1) = 0 Then Exit Sub
        DoEvents
        TheList.AddItem MyString$
    Wend
    Close #1
End Sub



Private Sub CommandBack_Click()
Picture1.Visible = False
End Sub

Private Sub Form_Load()

SetWindowPos FrmMain.hwnd, conHwndTopmost, 400, 400, 550, 355, conSwpNoActivate Or conSwpShowWindow
FrmMain.SetFocus

Command7_Click

End Sub

Public Function GetStrFromPtrA(ByVal lpszA As Long) As String

   GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
   Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
   
End Function



Private Sub List1_Click()
Text1.Text = List1.Text
End Sub

Private Sub Timer1_Timer()
If Text2.Text >= 1 Then
If Text3.Text + 1 = Val(Text2.Text) Then Command2_Click
End If

If Check3.Value = 1 Then
a = List1.ListCount - 1
Text1.Text = List1.Text
Call Run(Text1.Text)
If List1.ListIndex = a Then
List1.Selected(0) = True
Else
List1.Selected(List1.ListIndex + 1) = True
End If
Else
Call Run(Text1.Text)
End If

Text3.Text = Val(Text3.Text) + 1
End Sub

