Attribute VB_Name = "mod_DecodeKey"
'******************************************************************************************
'
'   DECODE a licence key
'
'   Created by Adrian Barnett
'   adriangbarnett@yahoo.co.uk
'   September 2009
'
'   This module needs to be in your target application. The end user will
'   pass his 64 character string to this the function ValLic()
'
'   To decode the Key, we need to rebuild the key from scratch.
'
'   It's Kind of brute force generating keys to get back in the application.
'   with a reduced amount of combinations because we know what were looking for.
'
'******************************************************************************************
Private MD5 As clsMD5        'dont remove this

Private LicKeyID As Long      'put this here incase we want to use the number in the ValLic() function  *optinal usage
Private ExpDate As String     'put this here so we can use it in the ValLic function                    *optinal usage



Public Function ValLic(appKey As String, appVersionNo As String, CustomerKey As String, ShowErrMsgBox As Boolean)
On Error GoTo err
'******************************************************************************************
'
'   Created by: adriangbarnett@yahoo.co.uk, September 2009
'
'   This function handle's the return codes from the Decoder.
'
'
'   ShowErrMsgBox    =  If set to true then the MsgBox "#00X#" will be displayed
'                       If it is set to False, only the 1 or 0 will be returned,( with no text boxes)
'
'
'       Returned Values from the Decoder.
'       0   DecodeLicKey()      Invalid Key
'       1   DecodeLicKey()      Critial Error (should never see this)
'       2   DecodeDateKey()     Critial error (should never see this)
'       3   DecodeDateKey()     Valid Key,Invalid Date
'       4   DecodeDateKey()     Valid Key Not Expired
'       5   DecodeDateKey()     Valid Key, Lic. but will expire after today
'       6   DecodeDateKey()     Valid Key, EXPIRED
'
'******************************************************************************************
    
    Dim RetValue As Integer         'The numeric value returned from the Decoder
    
    LicKeyID = 0                    'set as default Zero
    ExpDate = "???? / ?? / ??"      'Default zero value
    
    'Start to decode.
    RetValue = DecodeLicKey(appKey, appVersionNo, CustomerKey) 'Decode key
    
    'Make ExpDate readable
    ExpDate = Mid(ExpDate, 1, 4) & " / " & Mid(ExpDate, 5, 2) & " / " & Mid(ExpDate, 7, 2)
    
    
    'Perform actions beased onthe Decoder result.
    Select Case RetValue
        
        
        Case 0      'DecodeLicKey()      Invalid Key
                    
                    If ShowErrMsgBox = True Then MsgBox "#000# DecodeLicKey()" & vbCrLf & vbCrLf & "Invalid key", vbCritical, App.Title
                    ValLic = 1
                    Exit Function
             
            
        Case 1      'DecodeLicKey()     Critial Error (should never see this)
                    If ShowErrMsgBox = True Then MsgBox "#001# CRITICAL ERROR - DecodeLicKey()", vbCritical, App.Title
                    ValLic = 1
                    Exit Function
             
        
        Case 2      'DecodeDateKey()     Critial error (should never see this)
                    If ShowErrMsgBox = True Then MsgBox "#002# CRITICAL ERROR - DecodeDateKey()", vbExclamation, App.Title
                    ValLic = 1
                    Exit Function
            
            
        Case 3      'DecodeDateKey()     Valid Key, Invalid Date
                    If ShowErrMsgBox = True Then MsgBox "#003#" & vbCrLf & vbCrLf & "Invalid Date ID: " & LicKeyID, vbExclamation, App.Title
                    ValLic = 1
                    Exit Function
        
        
        Case 4      'DecodeDateKey()     Valid Key Not Expired
                    If ShowErrMsgBox = True Then MsgBox "#004#" & vbCrLf & vbCrLf & "Valid key ID: " & LicKeyID & vbCrLf & "Expires: " & ExpDate, vbInformation, App.Title
                    ValLic = 0
                    Exit Function
                    
                    
        Case 5      'DecodeDateKey()     Valid Key, but will expire after today
                    If ShowErrMsgBox = True Then MsgBox "#005#" & vbCrLf & vbCrLf & "Key expires today, ID: " & LicKeyID & vbCrLf & "Expires: " & ExpDate, vbExclamation, App.Title
                    ValLic = 0
                    Exit Function
                    
        
        Case 6      'DecodeDateKey()     Valid Key, EXPIRED
                    If ShowErrMsgBox = True Then MsgBox "#006# DecodeDateKey()" & vbCrLf & vbCrLf & "EXPIRED!, ID: " & LicKeyID & vbCrLf & "Expires: " & ExpDate, vbExclamation, App.Title
                    ValLic = 1
                    Exit Function
        
        Case Else   'unknown return value from Decoder.
                    If ShowErrMsgBox = True Then MsgBox "#00?# Unknown return value", vbCritical, App.Title
                    ValLic = 1
                    Exit Function
        
    End Select
    
    

'should never reach this point in the function
ValLic = 0  'return no error
Exit Function
err:
ValLic = 1  'return Error
End Function







Private Function DecodeLicKey(appKey As String, appVersionNo As String, CustomerKey As String)
On Error GoTo err

    '******************************************************************************************
    '   appKey          Custom Key String (ie Application Name)
    '   appVersionNo    app.major & app.minor
    '   CustomerKey     full 64 charcter string to decode
    '
    '       Returned Values from the Decoder.
    '       0   DecodeLicKey()      Invalid Key
    '       1   DecodeLicKey()      Critial Error (should never see this)
    '       2   DecodeDateKey()     Critial error (should never see this)
    '       3   DecodeDateKey()     Valid Key,Invalid Date
    '       4   DecodeDateKey()     Valid Key Not Expired
    '       5   DecodeDateKey()     Valid Key, Lic. but will expire after today
    '       6   DecodeDateKey()     Valid Key, EXPIRED
    '
    '******************************************************************************************
    
    Dim MAX_CUSTOMER_LIC_KEY_COUNT As Long  'Maximum number if customer serial keys, 10,000 is a good numbner, not to slow to do a check
    Dim i As Long                           'counter
    Dim KeyToHash As String                 'The string value to send to MD5.
    Dim KeyOutput As String                 'This is the result of the MD5 digest.
    Dim CustomerNumber As String            'a.k.a The Licence id counter. between 1 and 10,000
    Dim Ret As Integer                      'return Value

    Set MD5 = New clsMD5                    '<- dont remove this
    
    '**************************************************************************************************************
    'higher the number the longer it takes to decode, set default to 10,000, so u can have
    '10,000 unique licence keys, and can decode in about 3 seconds.
    MAX_CUSTOMER_LIC_KEY_COUNT = 10000
    '**************************************************************************************************************

            
    '**************************************************************************************************************
    '
    '   Here we need to recreate the MD5 Hash key, in the same way
    '   we Encoded it in the first place
    '
    '   Compare the 32 char hash key with the first 32 chars of the 64 char hash key supplied by customer
    '
    '   If the key is valid then a match should be found.
    '
    '   Once key match is found, we then check the Expiry date. (last 32 charcters of full serial key (CustomerKey)
    '
    '**************************************************************************************************************
    
                For i = 1 To MAX_CUSTOMER_LIC_KEY_COUNT 'i is the customer number, we need to find this in the hash key
                        DoEvents                        'between 1 and 10,000, start lets start searching.....
                        
                        'Merge stuff together
                        KeyToHash = appKey & appVersionNo & i

                        'Generate a MD5 hash of the key.
                        KeyOutput = MD5.CalculateMD5(KeyToHash)
      
                        'Compare if first 32 chars match
                        If UCase(KeyOutput) = Mid(UCase(CustomerKey), 1, 32) Then
                            'KEY MATCH FOUND

                            'Check Date Hash Key.
                                'i is the FOUND customer number
                                CustomerNumber = i
                                
                                'now lets decode the expire date
                                Ret = DecodeDateKey(appKey, appVersionNo, CustomerNumber, CustomerKey)
                                LicKeyID = i            'Set Current Customer Key ID, we might need this later
                                DecodeLicKey = Ret      'Return back to calling function
                                Exit Function
                        End If
                    Next


'No Key Match Found
DecodeLicKey = 0
Exit Function
err:
'MsgBox "Error DecodeLicKey() - " & err.Description, vbCritical, App.Title
DecodeLicKey = 1
End Function





Private Function DecodeDateKey(appKey As String, appVersionNo As String, CustomerNumber As String, CustomerKey As String)
On erorr GoTo err
    'We ony decode the date once the first part of licence key has been found\decoded.

    Dim DD As Integer                   'day (assume each month has 28 days, so we can avoid short\long months, leap year etc)
    Dim MM As Integer                   'month
    Dim YYYY As Integer                 'year
    Dim i As Long                       'counter
    Dim KeyToHash As String             'The string value to send to MD5.
    Dim KeyOutput As String             'This is the retsult of the MD5 digest.


    Dim tmpDateStr As String            'tmp string value of generating date...
    Dim tmpCurrentDateStr As String     'tmp string valuer of todays date


    Set MD5 = New clsMD5                '<- dont remove this
    
    
    For YYYY = Year(Date) - 5 To Year(Date) + 13   'calculate 3 years into the past and 10 - years in future for valid\expired keys
        DoEvents
        For MM = 1 To 12    'generate month 1-12
            DoEvents
            For DD = 1 To 28    'assume 28 days in every month, so we dont need to worry about februay, leap year and months with 30\31 days
                DoEvents
                '*******************************************************************
                '
                '   Build a string that conatins the generated date
                '
                '   Save into tmpDateStr, we need to compare the MD5 outrput with the user submitted key in a minute
                '
                '*******************************************************************
                
                tmpDateStr = YYYY
                'add month to string
                If MM < 10 Then
                    tmpDateStr = tmpDateStr & "0" & MM
                Else
                    tmpDateStr = tmpDateStr & MM
                End If
                
                'add day to string
                If DD < 10 Then
                    tmpDateStr = tmpDateStr & "0" & DD
                Else
                    tmpDateStr = tmpDateStr & DD
                End If
    
                    
                    'Merge stuff together
                    KeyToHash = appKey & appVersionNo & CustomerNumber & tmpDateStr

                    'Generate a MD5 hash of the key.
                    KeyOutput = MD5.CalculateMD5(KeyToHash)
    
                    'Compare if HASH key is a match with the last 32 charcters of the full serial number the end user submitted
                    If UCase(KeyOutput) = Mid(UCase(CustomerKey), 33, 32) Then
                            'DATE KEY MATCH FOUND
                            
                            
                            
                            'Save the found date into ExpDate, we mighht need this later in ValLic function
                            ExpDate = tmpDateStr
                            
                            
                            '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                            'Check if the key has expired, 1st we need to build todays date string
                            tmpCurrentDateStr = Year(Date)
                            
                            'Current MONTH
                            If Month(Date) < 10 Then
                                tmpCurrentDateStr = tmpCurrentDateStr & "0" & Month(Date)
                            Else
                                tmpCurrentDateStr = tmpCurrentDateStr & Month(Date)
                            End If
                            
                            'Current DAY
                            If Day(Date) < 10 Then
                                tmpCurrentDateStr = tmpCurrentDateStr & "0" & Day(Date)
                            Else
                                tmpCurrentDateStr = tmpCurrentDateStr & Day(Date)
                            End If
                            '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                            
                            '******************************************************************************************
                            '       Returned Values
                            '       2   DecodeDateKey()     Critial error (should never see this)
                            '       3   DecodeDateKey()     Valid Key,Invalid Date
                            '       4   DecodeDateKey()     Valid Key Not Expired
                            '       5   DecodeDateKey()     Valid Key, will expire after today
                            '       6   DecodeDateKey()     Valid Key, EXPIRED
                            '
                            '******************************************************************************************
                            
                            'Compare Date
                            If tmpCurrentDateStr < tmpDateStr Then
                                'Lic is still valid
                                DecodeDateKey = 4
                                Exit Function
                            End If
                            
                            'Compare Date, (todays date)
                            If tmpCurrentDateStr = tmpDateStr Then
                                'Lic. is still valid, but will expire after today
                                DecodeDateKey = 5
                                Exit Function
                            Else
                                'licence is expired
                                DecodeDateKey = 6
                                Exit Function
                            End If
 
                    End If
                    
    
            Next
        Next
    Next
    
    
'Customer Key is VALID but No Valid Date Key Found
DecodeDateKey = 3
Exit Function
err:
'MsgBox "Error DecodeDateKey() - " & err.Description
DecodeDateKey = 2
End Function
