Attribute VB_Name = "mod_EncodeKey"
'************************************************************************
'
'   ENCODE a licence key (create a serial key from strings and send to MD5 Hash
'
'   Created by Adrian Barnett
'   adriangbarnett@yahoo.co.uk
'   September 2009
'
'************************************************************************
Private MD5 As clsMD5

Public Function EncodeLicKey(appKey As String, appVersionNo As String, CustomerNumber As String, ExpireYYYY As String, ExpireMM As String, ExpireDD As String)
On Error GoTo err

    '
    '   appKey          Custom Key String (ie Application Name)
    '   appVersionNo    app.major & app.minor
    '   CustomerNumber  Unique number between 1 and MAX_NUMBER_OF_LICENCE keys (between 1 and 10,000)
    '   ExpireYYYY      Year of expiry date
    '   ExpireMM        Month of Expiry Date
    '   ExpireDD        Day of Expiry Date

    Dim KeyToHash As String
    Dim KeyOutput As String     'This is the retsult of the MD5 digest.
    
    Set MD5 = New clsMD5

    'Simple Error check data
    If Len(appKey) < 1 Then Exit Function
    If Len(appVersionNo) < 1 Then Exit Function
    If Len(CustomerNumber) < 1 Then Exit Function
    If Len(ExpireYYYY) < 1 Then Exit Function
    If Len(ExpireMM) < 1 Then Exit Function
    If Len(ExpireDD) < 1 Then Exit Function


    'Merge string into one
    KeyToHash = appKey & appVersionNo & CustomerNumber
    

    'Generate MD5 Digest of string
    KeyOutput = MD5.CalculateMD5(KeyToHash)
    
    'Returtn the full 32 charcter hash key
    EncodeLicKey = KeyOutput

Exit Function
err:
MsgBox "Error - EncodeLicKey() " & err.Description, vbCritical, App.Title
EncodeLicKey = ""
End Function

Public Function EncodeDateKey(appKey As String, appVersionNo As String, CustomerNumber As String, ExpireYYYY As String, ExpireMM As String, ExpireDD As String)
On Error GoTo err

    '
    '   appKey          Custom Key String (ie Application Name)
    '   appVersionNo    app.major & app.minor
    '   CustomerNumber  Unique number between 1 and MAX_NUMBER_OF_LICENCE keys you have created
    '   ExpireYYYY      Year of expiry date
    '   ExpireMM        Month of Expiry Date
    '   ExpireDD        Day of Expiry Date

    Dim KeyToHash As String
    Dim KeyOutput As String     'This is the retsult of the MD5 digest.
    
    Set MD5 = New clsMD5

    'Simple Error check data
    If Len(appKey) < 1 Then Exit Function
    If Len(appVersionNo) < 1 Then Exit Function
    If Len(CustomerNumber) < 1 Then Exit Function
    If Len(ExpireYYYY) < 1 Then Exit Function
    If Len(ExpireMM) < 1 Then Exit Function
    If Len(ExpireDD) < 1 Then Exit Function
    
    'Merge string into one, and create a Expire Date hash key
    KeyToHash = appKey & appVersionNo & CustomerNumber & ExpireYYYY & ExpireMM & ExpireDD
    

    'Generate MD5 Digest of string
    KeyOutput = MD5.CalculateMD5(KeyToHash)
    
    'Returtn the full 32 charcter hash key
    EncodeDateKey = KeyOutput

Exit Function
err:
MsgBox "Error - EncodeLicKey() " & err.Description, vbCritical, App.Title
EncodeDateKey = ""
End Function


