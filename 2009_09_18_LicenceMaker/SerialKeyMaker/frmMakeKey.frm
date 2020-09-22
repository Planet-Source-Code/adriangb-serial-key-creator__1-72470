VERSION 5.00
Begin VB.Form frmMakeKey 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Licence Key Maker - adriangbarnett@yahoo.co.uk"
   ClientHeight    =   9516
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9516
   ScaleWidth      =   9060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "Validate a Key"
      Height          =   852
      Left            =   240
      TabIndex        =   30
      Top             =   8520
      Width           =   8655
      Begin VB.CommandButton Command2 
         Caption         =   "Validate"
         Height          =   372
         Left            =   7680
         TabIndex        =   32
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtUserInputKey 
         Alignment       =   2  'Center
         Height          =   288
         Left            =   840
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   360
         Width           =   6735
      End
      Begin VB.Label Label16 
         Caption         =   "Test Key"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Key Results"
      Height          =   1452
      Left            =   240
      TabIndex        =   25
      Top             =   6840
      Width           =   8655
      Begin VB.TextBox txtFullKey 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   288
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "Text5"
         Top             =   1080
         Width           =   6735
      End
      Begin VB.TextBox txtKeyPublicKey1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   288
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Text5"
         Top             =   600
         Width           =   3732
      End
      Begin VB.TextBox txtKeyPublicKey2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   288
         Left            =   4200
         TabIndex        =   26
         Text            =   "Text5"
         Top             =   600
         Width           =   3732
      End
      Begin VB.Label Label17 
         Caption         =   "Key Part 1 (Customer ID Key \ a.k.a Licence Key ID)"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label13 
         Caption         =   "Key Part 2 (Expiry Date Key)"
         Height          =   255
         Left            =   4200
         TabIndex        =   34
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label14 
         Caption         =   "Full Serial Key"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Create Key"
      Height          =   732
      Left            =   240
      TabIndex        =   23
      Top             =   5880
      Width           =   8655
      Begin VB.CommandButton cmdCreateKey 
         Caption         =   "Create Serial Key"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   3360
         MaskColor       =   &H8000000F&
         TabIndex        =   24
         Top             =   240
         Width           =   2172
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Customer Number \ Licence Key Count\ ID"
      Height          =   1212
      Left            =   240
      TabIndex        =   19
      Top             =   4440
      Width           =   8655
      Begin VB.ComboBox cN 
         Height          =   315
         Left            =   5400
         TabIndex        =   20
         Text            =   "Combo1"
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label18 
         Caption         =   "Each N should be unique to a single end user."
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   840
         Width           =   4575
      End
      Begin VB.Label Label12 
         Caption         =   "N"
         Height          =   255
         Left            =   6000
         TabIndex        =   22
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Licence Key ID (Supports up to 10,000 user keys)"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Width           =   4575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Application Info"
      Height          =   2652
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   8775
      Begin VB.TextBox txtAppRev 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   288
         Left            =   5880
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   1560
         Width           =   852
      End
      Begin VB.TextBox txtAppMinor 
         Alignment       =   2  'Center
         Height          =   288
         Left            =   4800
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1560
         Width           =   852
      End
      Begin VB.TextBox txtAppMajor 
         Alignment       =   2  'Center
         Height          =   288
         Left            =   3720
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1560
         Width           =   852
      End
      Begin VB.TextBox txtappKey 
         Height          =   288
         Left            =   3720
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   360
         Width           =   3732
      End
      Begin VB.Label Label15 
         Caption         =   $"frmMakeKey.frx":0000
         Height          =   492
         Left            =   240
         TabIndex        =   29
         Top             =   2040
         Width           =   7932
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   7800
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label6 
         Caption         =   "Revision (We dont use this)"
         Height          =   252
         Left            =   5880
         TabIndex        =   17
         Top             =   1320
         Width           =   2172
      End
      Begin VB.Label Label5 
         Caption         =   "Minor"
         Height          =   252
         Left            =   4800
         TabIndex        =   15
         Top             =   1320
         Width           =   852
      End
      Begin VB.Label Label4 
         Caption         =   "Major"
         Height          =   252
         Left            =   3720
         TabIndex        =   13
         Top             =   1320
         Width           =   852
      End
      Begin VB.Label Label3 
         Caption         =   "Application Version Number"
         Height          =   252
         Left            =   240
         TabIndex        =   12
         Top             =   1560
         Width           =   2292
      End
      Begin VB.Label Label2 
         Caption         =   "This string needs to be hard coded into the application within the decode process."
         ForeColor       =   &H000000FF&
         Height          =   372
         Left            =   1440
         TabIndex        =   10
         Top             =   720
         Width           =   6252
      End
      Begin VB.Label Label1 
         Caption         =   "Custom Key String (ie Application Name)"
         Height          =   252
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   3252
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Expire Date"
      Height          =   1212
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   8775
      Begin VB.ComboBox cDD 
         Height          =   315
         Left            =   6720
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   600
         Width           =   612
      End
      Begin VB.ComboBox cMM 
         Height          =   315
         Left            =   5400
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   600
         Width           =   612
      End
      Begin VB.ComboBox cYYYY 
         Height          =   315
         Left            =   3960
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   600
         Width           =   972
      End
      Begin VB.Label Label10 
         Caption         =   "The licence will become invalid beyond this date. The decoder supports -3 years and +10 years."
         Height          =   615
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label9 
         Caption         =   "DD"
         Height          =   255
         Left            =   6840
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "MM"
         Height          =   255
         Left            =   5520
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "YYYY"
         Height          =   255
         Left            =   4200
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmMakeKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************
'
'   ENCODE\DECODE a licence key (create a licence key from strings and send to MD5 Hash
'
'   Created by Adrian Barnett
'   adriangbarnett@yahoo.co.uk
'   September 2009
'
'************************************************************************


Dim YYYY() As String
Dim DD() As String
Dim MM() As String
Dim NN() As String




Private Sub cDD_Change()
Call ChkData
End Sub

Private Sub cDD_LostFocus()
Call ChkData
End Sub

Private Sub cmdCreateKey_Click()

    Dim a As String
    Dim b As String
    
    
    With Me
    
        'disable, so use does not click button two times.
        .cmdCreateKey.Enabled = False
    
        'Make a licence key
        a = EncodeLicKey(.txtappKey.Text, .txtAppMajor.Text & .txtAppMinor.Text, .cN.Text, .cYYYY.Text, .cMM.Text, .cDD.Text)
        b = EncodeDateKey(.txtappKey.Text, .txtAppMajor.Text & .txtAppMinor.Text, .cN.Text, .cYYYY.Text, .cMM.Text, .cDD.Text)
        
        'set txt field
        .txtKeyPublicKey1 = UCase(a)
        .txtKeyPublicKey2 = UCase(b)
        
        'combined, FULL SERIAL KEY (give to end user)
        .txtFullKey.Text = UCase(a) & UCase(b)
        
        'done
        .cmdCreateKey.Enabled = True
        
    
    End With
    


End Sub

Private Sub cMM_Change()
Call ChkData
End Sub

Private Sub cMM_LostFocus()
Call ChkData
End Sub

Private Sub cN_Change()
Call ChkData
End Sub

Private Sub cN_LostFocus()
Call ChkData
End Sub

Private Sub Command2_Click()

    
    Me.Command2.Enabled = False
    Me.Command2.Caption = "Wait..."
        
        'Validate a licence key
        Call ValLic(Me.txtappKey.Text, Me.txtAppMajor.Text & Me.txtAppMinor.Text, Me.txtUserInputKey.Text, True)
        
    Me.Command2.Caption = "Validate"
    Me.Command2.Enabled = True

End Sub

Private Sub cYYYY_Change()
Call ChkData
End Sub

Private Sub cYYYY_LostFocus()
Call ChkData
End Sub

Private Sub Form_Load()


    'seta form defaults
    Dim i As Long
    ReDim YYYY(0)
    ReDim DD(0)
    ReDim MM(0)
    ReDim NN(0)
    
    With Me
    
        'set some defaults
        Me.txtappKey.Text = App.Title
        Me.txtAppMajor.Text = App.Major
        Me.txtAppMinor.Text = App.Minor
        Me.txtAppRev.Text = "N/A"
        
        Me.txtKeyPublicKey1.Text = ""
        Me.txtKeyPublicKey2.Text = ""
        Me.txtFullKey.Text = ""

        
        Me.txtUserInputKey.Text = ""
    
        With .cN
            .Clear
            For i = 1 To 10000
                .AddItem i
                ReDim Preserve NN(UBound(NN) + 1)
                NN(UBound(NN)) = i
            Next
                .Text = "1"
        End With
    
    
        'Year
        With .cYYYY
            .Clear
            For i = -1 To 10
                .AddItem Year(Date) + i
                ReDim Preserve YYYY(UBound(YYYY) + 1)
                YYYY(UBound(YYYY)) = Year(Date) + i
                
            Next
        End With
        
        'Month
        With .cMM
            .Clear
            For i = 1 To 12
                If i < 10 Then
                    .AddItem "0" & i
                    ReDim Preserve MM(UBound(MM) + 1)
                    MM(UBound(MM)) = "0" & i
                Else
                    .AddItem i
                    ReDim Preserve MM(UBound(MM) + 1)
                    MM(UBound(MM)) = i
                End If
            Next
        End With

        'Day
        With .cDD
            .Clear
            For i = 1 To 28 'so user  dont acceintly select 31st Februay
                If i < 10 Then
                    .AddItem "0" & i
                    ReDim Preserve DD(UBound(DD) + 1)
                    DD(UBound(DD)) = "0" & i
                Else
                    .AddItem i
                    ReDim Preserve DD(UBound(DD) + 1)
                    DD(UBound(DD)) = i
                End If
            Next
        End With


        'set current date
        .cYYYY = Year(Date)
        
        'Month
        If Len(Month(Date)) < 2 Then
            .cMM = "0" & Month(Date)
        Else
            .cMM = Month(Date)
        End If
        
        'Day
        If Len(Day(Date)) < 2 Then
            .cDD = "0" & Day(Date)
        Else
           .cDD = Day(Date)
        End If

    End With
    

End Sub


'check data is all valid before generate button is pressed.
Public Function ChkData()
On Error GoTo err

    Dim s As String
    Dim eFlag As Boolean
    
    
    eFlag = False

    If IsInArr(NN, Me.cN.Text) <> 0 Then eFlag = True          'chk selected Customer number valid in array
    If IsInArr(YYYY, cYYYY.Text) <> 0 Then eFlag = True      'chk selected year number valid in array
    If IsInArr(MM, cMM.Text) <> 0 Then eFlag = True          'chk selected month number valid in array
    If IsInArr(DD, cDD.Text) <> 0 Then eFlag = True          'chk selected Day number valid in array
    If Len(Me.txtappKey.Text) < 1 Then eFlag = True        'appKey
    If Len(Me.txtAppMajor.Text) < 1 Then eFlag = True       'Major version
    If Len(Me.txtAppMinor.Text) < 1 Then eFlag = True       'minor
        
        
        
    If Len(cN) < 1 Then eFlag = True
    If Len(cYYYY) < 1 Then eFlag = True
    If Len(cMM) < 1 Then eFlag = True
    If Len(cDD) < 1 Then eFlag = True

    
    If eFlag = True Then
        cmdCreateKey.Enabled = False
    Else
        cmdCreateKey.Enabled = True
    End If
    
    
Exit Function
err:
End Function


Private Function IsInArr(ArrayName As Variant, StrValue As String)
On Error GoTo err

    'check select string exists inside an array
    Dim i As Integer
    
    For i = 1 To UBound(ArrayName)
    
        If StrValue = ArrayName(i) Then
            'match
            IsInArr = 0
            Exit Function
        End If

    Next

    'Not found in array
    IsInArr = 1
    Exit Function
    
    
err:
IsInArr = 1
End Function

Private Sub txtappKey_Change()
Call ChkData
End Sub

Private Sub txtAppMajor_Change()
Call ChkData
End Sub

Private Sub txtAppMinor_Change()
Call ChkData
End Sub

Private Sub txtAppRev_Change()
Call ChkData
End Sub
