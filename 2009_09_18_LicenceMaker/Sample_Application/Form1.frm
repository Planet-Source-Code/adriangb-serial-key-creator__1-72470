VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2304
   ClientLeft      =   96
   ClientTop       =   396
   ClientWidth     =   9276
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2304
   ScaleWidth      =   9276
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEndUserKey 
      Height          =   288
      Left            =   2640
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1320
      Width           =   6375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00000000&
      Height          =   288
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   720
      Width           =   732
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00000000&
      Height          =   288
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   720
      Width           =   852
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   2640
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   3612
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test my key"
      Height          =   372
      Left            =   2880
      TabIndex        =   0
      Top             =   1800
      Width           =   4212
   End
   Begin VB.Label Label3 
      Caption         =   "Enter a Valid Key"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Version (Major \ Minor)"
      Height          =   252
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   2172
   End
   Begin VB.Label Label1 
      Caption         =   "Application Name\Key"
      Height          =   252
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2292
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Dim Ret As Integer
    

    Me.Command1.Enabled = False
    Me.Command1.Caption = "Checking, Please wait...."
    
    'Decode the key the user entered
    Ret = ValLic(Text1.Text, App.Major & App.Minor, Me.txtEndUserKey.Text, False)
    
    Me.Command1.Caption = "Test my key"
    Me.Command1.Enabled = True
    
    
    If Ret = 1 Then
        MsgBox "Key is invalid or expired", vbCritical
        Exit Sub
    End If
    
    If Ret = 0 Then
        MsgBox "Key is valid", vbInformation
        Exit Sub
    End If
    
    
    'should never get to this poin in the application
    MsgBox "Unable to Resolve resturn value " & Ret, vbCritical
    

End Sub

Private Sub Form_Load()


    With Me
        .Text1.Text = App.Title
        .Text2.Text = App.Major
        .Text3.Text = App.Minor
        .txtEndUserKey.Text = "???????????????"
    End With
    
End Sub
