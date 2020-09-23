Attribute VB_Name = "modStrings"
Option Explicit

Public Function AddFile(ByVal strDirectory As String, _
                        ByVal strFileName As String) _
                        As String
    'This will add a file name to a directory path to
    'create a full filepath.
    
    If Right(strDirectory, 1) <> "\" Then
        'insert a backslash
        strDirectory = strDirectory & "\"
    End If
    
    'append the file name to the directory path now
    AddFile = strDirectory & strFileName
End Function

Public Function CommaCount(ByVal strLine As String) _
                           As Integer
    'This will return the number of commas foun in the string. Mainly
    'use to find the number of variables declared on the same line
    
    Dim intCounter As Integer
    Dim intLastPos As Integer
    Dim intCommaNum As Integer
    
    intLastPos = 0
    
    Do
        intCounter = InStr(intLastPos + 1, strLine, ",")
        
        If intCounter <> 0 Then
            intCommaNum = intCommaNum + 1
        End If
        intLastPos = intCounter
    Loop Until intLastPos = 0
    
    'return result
    CommaCount = intCommaNum
End Function

Public Function GetAfter(ByVal strSentence As String, _
                         Optional ByVal strCharacter As String = "=") _
                         As String
    'This procedure returns all the character of a
    'string after the "=" sign.
    
    Dim intCounter As Integer
    Dim strRest As String
    Dim strSign As String
    
    strSign = strCharacter
    
    'find the last position of the specified sign
    intCounter = InStrRev(strSentence, strSign)
    
    If intCounter <> Len(strSentence) Then
        strRest = Right(strSentence, (Len(strSentence) - (intCounter + Len(strCharacter) - 1)))
    Else
        strRest = ""
    End If
    
    GetAfter = strRest
End Function

Public Function GetBefore(ByVal strSentence As String) _
                          As String
    'This procedure returns all the character of a
    'string before the "=" sign.
    
    Const Sign = "="
    
    Dim intCounter As Integer
    Dim strBefore As String
    
    'find the position of the equals sign
    intCounter = InStr(1, strSentence, Sign)
    
    If (intCounter <> Len(strSentence)) And (intCounter <> 0) Then
        strBefore = Left(strSentence, (intCounter - 1))
    Else
        strBefore = ""
    End If
    
    GetBefore = strBefore
End Function

Public Sub GetFileList(ByRef strFiles() As String, _
                       Optional ByVal strPath As String, _
                       Optional ByVal strExtention As String = "*.*", _
                       Optional ByVal lngAttributes As Long = vbNormal, _
                       Optional ByVal intNumFiles As Integer)
    'This procedure will get a list of files
    'available in the specified directory. If
    'no directory is specified, then the
    'applications directory is taken to be
    'the default.
    
    Dim intCounter As Integer       'used to reference new elements in the array
    Dim strTempName As String       'temperorily holds a file name
    
    'validate the parameters for correct values
    If (Trim(strPath = "")) _
       Or (Dir(strPath, vbDirectory) = "") Then
        
        'invalid path, assume applications
        'directory
        strPath = App.Path
    End If
    
    'reset the array before entering new data
    ReDim strFiles(0)
    
    'resize the array to nothing if the
    'number of files specified is less
    'than can be returned
    If intNumFiles < 1 Then
        'return the maximum number of files (if possible)
        intNumFiles = 32767
    End If
    
    'include a wild card if the user only
    'specified the extention
    If Left(strExtention, 1) = "." Then
        strExtention = "*" & strExtention
    ElseIf InStr(strExtention, ".") = 0 Then
        strExtention = "*." & strExtention
    End If
    
    'get the first file name to start
    'the file search for this directory
    strTempName = Dir(AddFile(strPath, _
                              strExtention), _
                      lngAttributes)
    
    'keep getting new files until there are
    'no more to return
    Do While (strTempName <> "") _
       And (intCounter <= intNumFiles)
        
        'enter the file into the array
        ReDim Preserve strFiles(intCounter)
        strFiles(intCounter) = strTempName
        intCounter = intCounter + 1
        
        'get a new file
        strTempName = Dir
    Loop
End Sub

Public Function IsNotInQuote(ByVal strText As String, _
                             ByVal strWords As String) _
                             As Boolean
    'This function will tell you if the specified text is in quotes within
    'the second string. It does this by counting the number of quotation
    'marks before the specified strWords. If the number is even, then the
    'strWords are not in qototes, otherwise they are.
    
    'the quotation mark, " , is ASCII character 34
    
    Dim lngGotPos As Long
    Dim lngCounter As Long
    Dim lngNextPos As Long
    
    'find where the position of strWords in strText
    lngGotPos = InStr(1, strText, strWords)
    If lngGotPos = 0 Then
        IsNotInQuote = True
        Exit Function
    End If
    
    'start counting the number of quotation marks
    lngNextPos = 0
    Do
        lngNextPos = InStr(lngNextPos + 1, strText, Chr(34))
        
        If (lngNextPos <> 0) And (lngNextPos < lngGotPos) Then
            'quote found, add to total
            lngCounter = lngCounter + 1
        End If
    Loop Until (lngNextPos = 0) Or (lngNextPos >= lngGotPos)
    
    'no quotes at all found
    If lngCounter = 0 Then
        IsNotInQuote = True
        Exit Function
    End If
    
    'if the number of quotes is even, then return true, else return false
    If lngCounter Mod 2 = 0 Then
        IsNotInQuote = True
    End If
End Function

Public Function GetPath(ByVal strAddress As String) _
                        As String
    'This function returns the path from a string containing the full
    'path and filename of a file.
    
    Dim intLastPos As Integer
    
    'find the position of the last "\" mark in the string
    intLastPos = InStrRev(strAddress, "\")
    
    'if no \ found in the string, then
    If intLastPos = 0 Then
        'return the whole string
        intLastPos = Len(strAddress) + 1
    End If
    
    'return everything before the last "\" mark
    GetPath = Left(strAddress, (intLastPos - 1))
End Function

Public Function IsWord(ByVal strLine As String, _
                       ByVal strWord As String) _
                       As Boolean
    'This function will return True if the
    'specified word is not part of another
    'word
    
    Dim blnLeftOk As Boolean    'the left side of the word is valid
    Dim blnRightOk As Boolean   'the right side of the word is valid
    Dim lngWordPos As Long      'the position of the specified word in the string
    
    If (Len(strWord) > Len(strLine)) _
       Or (strLine = "") _
       Or (strWord = "") Then
        'invalid parameters
        IsWord = False
        Exit Function
    End If
    
    'remove leading/trailing spaces
    strLine = Trim(strLine)
    strWord = Trim(strWord)
    
    lngWordPos = InStr(UCase(strLine), UCase(strWord))
    
    If lngWordPos = 0 Then
        'word not found
        IsWord = False
        Exit Function
    End If
    
    'check the left side of the word
    If lngWordPos = 1 Then
        'word is on the left side of the string
        blnLeftOk = True
    Else
        'check the character to the left of the word
        Select Case UCase(Mid(strLine, lngWordPos - 1, 1))
        Case "A" To "Z", "0" To "9"
        Case Else
            blnLeftOk = True
        End Select
    End If
    
    'check the right side of the word
    If (lngWordPos + Len(strWord)) = Len(strLine) Then
        'word is on the left side of the string
        blnRightOk = True
    Else
        'check the character to the left of the word
        'Debug.Print strWord, UCase(Mid(strLine, lngWordPos + Len(strWord), 1))
        Select Case UCase(Mid(strLine, lngWordPos + Len(strWord), 1))
        Case "A" To "Z", "0" To "9"
            'Stop
        Case Else
            blnRightOk = True
        End Select
    End If
    
    'if both sides are OK, then return True
    If blnLeftOk And blnRightOk Then
        IsWord = True
    End If
End Function

Public Function PaddString(ByVal strText As String, _
                           ByVal lngTotalChar As Long) _
                           As String
    'This will padd a string with trailing spaces so
    'that the returned string matches the total
    'number of characters specified. If the string
    'passed is bigger than the total number of
    'characters, then the string is truncated and then
    'returned.
    
    Dim lngLenText As Long  'the length of the text string passed
    
    'if the number of characters is zero, then
    'return nothing
    If lngTotalChar = 0 Then
        PaddString = ""
        Exit Function
    End If
    
    'get the length of the string
    lngLenText = Len(strText)
    
    If lngLenText >= lngTotalChar Then
        'return a trucated string
        PaddString = Left(strText, lngTotalChar)
    Else
        'padd the string out with spaces
        PaddString = strText & Space(lngTotalChar - lngLenText)
    End If
End Function

Public Function StripQuotes(ByVal strText As String) As String
    'This function will remove all text between the
    'quotation marks (")
    
    Const QUOTE = """"         'the quotation mark (")
    
    Dim lngQuoteStart As Long       'the position of the first quotation mark found in the string
    Dim lngQuoteFinish As Long      'the position of the quote mark after the first position
    
    'get the position of a quotation mark
    lngQuoteStart = InStr(strText, QUOTE)
    
    Do While (lngQuoteStart > 0)
        'find the next quote mark after the found position
        lngQuoteFinish = InStr(lngQuoteStart + 1, _
                               strText, _
                               QUOTE)
        
        'if a second quotation mark was found, remove
        'all text between
        If lngQuoteFinish > 0 Then
            strText = Left(strText, _
                           lngQuoteStart - 1) & _
                      Right(strText, _
                            Len(strText) - lngQuoteFinish)
        Else
            'there is only one quotation mark. Remove it
            strText = Left(strText, _
                           lngQuoteStart - 1) & _
                      Right(strText, _
                            Len(strText) - lngQuoteStart)
        End If
        
        'get the next occurance of a quotation mark
        lngQuoteStart = InStr(lngQuoteStart, _
                              strText, _
                              QUOTE)
    Loop
    
    'return the stripped text
    StripQuotes = strText
End Function
