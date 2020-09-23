Attribute VB_Name = "modGeneral"
Option Explicit

' The object declaration
Public g_objFileSystem As Scripting.FileSystemObject
Public g_objTextStream As Scripting.TextStream

' File name wher the information is stored.
Public Const g_sFileName = "FileSystemDemo.txt"

Public Enum SeperateInfo
    RightSide = 1
    LeftSide = 2
End Enum

' Just to make the coding easy declare these enum
Public Enum FileMode
    ReadMode = 1
    AppendMode = 2
    WriteMode = 3
End Enum

' Topic Info
Public Type TopicInfo
    bCodeSampleAvailable As Boolean
    sTopicTitle As String
    sTopicDescription As String
End Type

Public g_arrTopics() As TopicInfo
Public g__iNumTopics As Integer


Public Function OpenTextFile(sFileName As String, iMode As FileMode) As Boolean
' This function will open the text file using the OpenTextfile Method.

On Error GoTo ErrOpen

' iMode = 1 means open for Read ; iMode = 2 means open for Append ; iMode = 3 means open for write
' Open file for Read or Append
    
    Select Case iMode
        
        Case 1
            Set g_objTextStream = g_objFileSystem.OpenTextFile(sFileName, ForReading, False, TristateFalse)
        Case 2
            Set g_objTextStream = g_objFileSystem.OpenTextFile(sFileName, ForAppending)
        Case 3
            Set g_objTextStream = g_objFileSystem.OpenTextFile(sFileName, ForWriting)
    End Select
    
    OpenTextFile = True
    Exit Function
    
ErrOpen:

    Debug.Print "Err Open file : " + Err.Description
    
End Function


Public Function ReadAndFillArray() As Boolean

' Read the file and fill the data in the array

On Error GoTo ErrRead
Dim sFileName As String
Dim sFromFile As String
Dim bStart As Boolean

Dim bTopicFirstLine As Boolean
Dim sDescriptionText As String

Dim iTopicCount As Integer      ' keep track the Topic title information
Const Topic = "#Topic"
    
    sFileName = App.Path + "\" + g_sFileName
    
    ' Check if the file exists
    If Dir(sFileName, vbNormal) = "" Then
        MsgBox "Source file : " + sFileName + " doesnot exist. "
        Exit Function
    End If
    
    If Not OpenTextFile(sFileName, ReadMode) Then
        MsgBox "Cannot open the file : " + sFileName
        Exit Function
    End If
        
    ' Till we get the end of file or #End mark in the file, read the file
    ' Skip the description portion
    
    Do While Not g_objTextStream.AtEndOfStream
        
        ' Read the line from the text file
        sFromFile = g_objTextStream.ReadLine
        
        Debug.Print sFromFile       ' Just to view the information in debug window
        
        ' Just check the line before entering into the fn
        'If LCase(sFromFile) = LCase("#Start") Then bStart = True
        
        If bStart = False Then  ' Just leave the comment portion
            
            If LCase(sFromFile) = LCase("#Start") Then bStart = True
            
            If bStart Then  ' Line start has appeared in the file
            
                ' Get the topics count and set to the number of topics variable
                g__iNumTopics = CInt(g_objTextStream.ReadLine)
                
                ' Redim the array to hold the content from file
                If g__iNumTopics = 0 Then
                    
                    MsgBox "No topics found "
                    Exit Function
                
                Else        ' g_iNumTopics = 0 Else
                
                    ReDim Preserve g_arrTopics(g__iNumTopics)
                    ' set the flag to start reading the topic info
                    sDescriptionText = ""
                
                End If  ' g_iNumTopics = 0 End if
            
            End If  ' bStart End if
            
        Else
            
            ' Here we have to track the start of topic, first line of topic and skip line
            If Left(LCase(sFromFile), 9) = LCase("#SkipLine") Then GoTo LineSkipped
            
            ' Is the line reperesent the next topic start
            If sFromFile = Topic + Trim$(Str(iTopicCount + 1)) Then
                
                iTopicCount = iTopicCount + 1
                
                ' Get the first line from the topic. Get the topic and code sample available flag.
                
                ' Before refreshing the Description Text put that to the previous array element's info
                If iTopicCount > 1 Then     ' For the first item we won't have the previous item.
                    g_arrTopics(iTopicCount - 1).sTopicDescription = sDescriptionText
                End If
                
                sDescriptionText = ""
                
                sFromFile = g_objTextStream.ReadLine
                
                ' If the char "|" found in the line then get the topic title and code sample available flag.
                If CharFound(sFromFile, "|") Then
                    g_arrTopics(iTopicCount).sTopicTitle = SeperateString(sFromFile, "|", LeftSide)
                    g_arrTopics(iTopicCount).bCodeSampleAvailable = IIf(CBool(SeperateString(sFromFile, "|", RightSide)) = True, True, False)
                Else
                    g_arrTopics(iTopicCount).sTopicTitle = sFromFile
                End If
            
            Else        ' The line doesnot represent any topic
                
                ' If not the end of file info then add the string to the description
                
                If LCase(sFromFile) = LCase("#End") Then
                    ' We don't want any information afterwards
                    g_arrTopics(iTopicCount).sTopicDescription = sDescriptionText
                    
                    ' We got the needed data from the file. Exit
                    ReadAndFillArray = True
                    Exit Function
                
                Else    ' If not end then keep on add the content to the description text
                    
                    If sDescriptionText = "" Then
                        sDescriptionText = sFromFile
                    Else
                        sDescriptionText = sDescriptionText + vbCrLf + sFromFile
                    End If
                End If
                
            End If ' End if of sFromFile = any topic
            
LineSkipped:

        End If
        
    Loop
    
    ReadAndFillArray = True

ErrRead:

End Function

'#SkipLine
'#End


Public Function CharFound(ByVal sSearchString As String, ByVal sSearchChar As String) As Boolean
' This function will return True if the Specified character is found in the string
Dim charpos As Integer
On Error GoTo ErrHandler

    charpos = InStr(1, sSearchString, sSearchChar, vbTextCompare)
    
    If charpos > 0 Then
        CharFound = True
    Else
        CharFound = False
    End If
    
Exit Function

ErrHandler:
    CharFound = False
End Function

Public Function SeperateString(ByVal sInText As String, ByVal Pattern As String, ByVal RightOrLeft As SeperateInfo)

On Error GoTo ErrSep
' This function will extract the  string to the left or right of the pattern string
    
    Dim DotPos As Integer
    Dim Loc As Boolean
    Dim TextLength As Integer

    TextLength = Len(sInText)
    
    Select Case UCase(RightOrLeft)
        Case 2      ' Left
            
            DotPos = InStr(1, sInText, Pattern, vbTextCompare)
            SeperateString = ""
            If DotPos > 0 Then
                SeperateString = Left(sInText, DotPos - 1)
            Else    ' Return blank string
                SeperateString = ""
                'SeperateString = sInText
            End If
            
        Case 1  ' Right
            
            DotPos = InStr(1, sInText, Pattern, vbTextCompare)
            SeperateString = ""
            If DotPos > 0 Then
                SeperateString = Right(sInText, TextLength - (DotPos + (Len(Pattern) - 1)))
            Else    ' Return blank string
                SeperateString = ""
                'SeperateString = sInText
            End If
    
    End Select
    
ErrSep:

End Function


Public Function GetFileName(sFileName As String) As String
' Get the file name
On Error GoTo ErrGet
Dim arrTmp() As String

Dim sTemp As String
Dim iLastElement As Integer

    If sFileName = "" Then Exit Function

    arrTmp = Split(sFileName, "\")
    iLastElement = UBound(arrTmp)
    
    sTemp = arrTmp(iLastElement)
    
    GetFileName = sTemp
    
    Exit Function

ErrGet:
    Debug.Print "Get File Name : " + Err.Description
End Function

