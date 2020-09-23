VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remove IE Ratings Password "
   ClientHeight    =   1950
   ClientLeft      =   3705
   ClientTop       =   3180
   ClientWidth     =   4365
   Icon            =   "Remove IE Ratings Password.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4365
   Begin VB.CommandButton cmdRemovePassword 
      Caption         =   "Remove Ratings &Password"
      Default         =   -1  'True
      Height          =   465
      Left            =   870
      TabIndex        =   0
      Top             =   900
      Width           =   2685
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   1650
      Width           =   4365
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "click on the button below..."
      Height          =   195
      Left            =   390
      TabIndex        =   2
      Top             =   510
      Width           =   1920
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "To remove the Internet Explorer Ratings Password,"
      Height          =   195
      Left            =   390
      TabIndex        =   1
      Top             =   300
      Width           =   3600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strTemp As String

Private Sub cmdRemovePassword_Click()
  If (strTemp <> "Error") And (strTemp <> Chr(0)) Then
    strTemp = Chr(0)
    SetBinaryValue "HKEY_LOCAL_MACHINE\" & _
            "SOFTWARE\" & _
            "MICROSOFT\" & _
            "WINDOWS\" & _
            "CURRENTVERSION\" & _
            "POLICIES\" & _
            "RATINGS", _
            "Key", strTemp
    LookItUp
  Else
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  LookItUp
End Sub

Sub LookItUp()

  strTemp = GetBinaryValue("HKEY_LOCAL_MACHINE\" & _
            "SOFTWARE\" & _
            "MICROSOFT\" & _
            "WINDOWS\" & _
            "CURRENTVERSION\" & _
            "POLICIES\" & _
            "RATINGS", _
            "Key")
    
  If (strTemp <> "Error") And (strTemp <> Chr(0)) Then
    lblStatus.Caption = "Ratings Password Exists"
    cmdRemovePassword.Caption = "Remove Ratings &Password"
  Else
    lblStatus.Caption = "No Ratings Password Found..."
    cmdRemovePassword.Caption = "E &x i t"
  End If

End Sub
