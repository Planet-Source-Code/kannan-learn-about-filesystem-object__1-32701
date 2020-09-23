VERSION 5.00
Begin VB.Form frmSample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sample..."
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10920
   Icon            =   "frmSampleCode.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   10920
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraNoSample 
      Height          =   900
      Left            =   8955
      TabIndex        =   23
      Top             =   345
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Sorry !. No sample available for this topic now."
         Height          =   465
         Left            =   120
         TabIndex        =   24
         Top             =   270
         Width           =   2430
      End
   End
   Begin VB.Frame fraSkipLine 
      Caption         =   "Skip Line:"
      Height          =   2160
      Left            =   4590
      TabIndex        =   14
      Top             =   3120
      Visible         =   0   'False
      Width           =   4245
      Begin VB.TextBox txtSkipLine 
         Height          =   1275
         Left            =   75
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   22
         Top             =   195
         Width           =   4080
      End
      Begin VB.CommandButton cmdSkipLine 
         Caption         =   "SkipLine"
         Height          =   390
         Left            =   3120
         TabIndex        =   20
         Top             =   1635
         Width           =   900
      End
      Begin VB.Label Label7 
         Caption         =   "Skips the odd/even lines from the samplefile.txt file and display the remaining content"
         Height          =   645
         Left            =   180
         TabIndex        =   21
         Top             =   1500
         Width           =   2805
      End
   End
   Begin VB.Frame fraRead 
      Caption         =   "Read :"
      Height          =   2670
      Left            =   4575
      TabIndex        =   3
      Top             =   330
      Visible         =   0   'False
      Width           =   4245
      Begin VB.TextBox txtRead 
         Height          =   1200
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   18
         Top             =   765
         Width           =   4005
      End
      Begin VB.CommandButton cmdRead 
         Caption         =   "Read"
         Height          =   390
         Left            =   3120
         TabIndex        =   17
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtReadChars 
         Height          =   345
         Left            =   2880
         TabIndex        =   15
         Top             =   225
         Width           =   1200
      End
      Begin VB.Label Label6 
         Caption         =   "Reads the n from the file samplefile.txt"
         Height          =   300
         Left            =   90
         TabIndex        =   19
         Top             =   2235
         Width           =   2850
      End
      Begin VB.Label Label5 
         Caption         =   "Enter the number of chars to be read"
         Height          =   300
         Left            =   90
         TabIndex        =   16
         Top             =   270
         Width           =   2640
      End
   End
   Begin VB.Frame fraReadLine 
      Caption         =   "ReadLine :"
      Height          =   1905
      Left            =   195
      TabIndex        =   2
      Top             =   5355
      Visible         =   0   'False
      Width           =   4245
      Begin VB.TextBox txtReadLine 
         Height          =   1020
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   12
         Top             =   240
         Width           =   4095
      End
      Begin VB.CommandButton cmdReadLine 
         Caption         =   "Read Line"
         Height          =   390
         Left            =   3030
         TabIndex        =   11
         Top             =   1410
         Width           =   1110
      End
      Begin VB.Label Label4 
         Caption         =   "Reads line by line from the file app.path + ""\"" + samplefile.txt and displays here."
         Height          =   450
         Left            =   60
         TabIndex        =   13
         Top             =   1335
         Width           =   2910
      End
   End
   Begin VB.Frame fraReadAll 
      Caption         =   "ReadAll :"
      Height          =   3645
      Left            =   195
      TabIndex        =   1
      Top             =   1650
      Visible         =   0   'False
      Width           =   4245
      Begin VB.TextBox txtReadAll 
         Height          =   2730
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         Top             =   195
         Width           =   4065
      End
      Begin VB.CommandButton cmdReadAll 
         Caption         =   "Read All"
         Height          =   390
         Left            =   3105
         TabIndex        =   8
         Top             =   3075
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "Reads the file app.path + ""\"" + samplefile.txt and fills the textbox"
         Height          =   510
         Left            =   120
         TabIndex        =   9
         Top             =   3015
         Width           =   2835
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraCreateFile 
      Caption         =   "CreateFile :"
      Height          =   1335
      Left            =   195
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   4260
      Begin VB.TextBox txtCreatefile 
         Height          =   285
         Left            =   1470
         TabIndex        =   5
         Text            =   "a.txt"
         Top             =   345
         Width           =   2655
      End
      Begin VB.CommandButton cmdCreateFile 
         Caption         =   "Create "
         Height          =   390
         Left            =   2745
         TabIndex        =   4
         Top             =   795
         Width           =   1365
      End
      Begin VB.Label Label2 
         Caption         =   "File with the name app.path + ""\"" + textbox.text will be created."
         Height          =   465
         Left            =   135
         TabIndex        =   7
         Top             =   735
         Width           =   2505
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Enter File Name:"
         Height          =   240
         Left            =   150
         TabIndex        =   6
         Top             =   375
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frmSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_iCurScreen As Integer    ' will be set from the caller
Private m_objFS As Scripting.FileSystemObject
Private m_objTextStream As Scripting.TextStream
Private m_objFile As Scripting.File ' To access a file from the harddisk

Const Con_lExtraHt = 500
Const Con_lExtraWt = 150
Const ConSampleFile = "SampleFile.txt"      ' sample file name used

Public Property Let CurrentTopic(vData As Integer)
' This will be set by the caller
    m_iCurScreen = vData
    SetCurScreen
End Property

Private Sub SetCurScreen()

' Show the proper layer for the current example selected
    'HideAllLayers
    
    Select Case m_iCurScreen
    
        Case 2      ' CreateFile
            ShowLayer fraCreateFile
        Case 4      'Read
            ShowLayer fraRead
        Case 5      'ReadAll
            ShowLayer fraReadAll
        Case 6      ' ReadLine
            ShowLayer fraReadLine
        Case 8      ' SkipLine
            ShowLayer fraSkipLine
        Case Else
            ShowLayer fraNoSample
    End Select

End Sub

Private Sub HideAllLayers()

' Hide all the layers

On Error Resume Next

Dim bBool As Boolean

    bBool = False
    
    fraCreateFile.Visible = bBool
    fraNoSample.Visible = bBool
    fraRead.Visible = bBool
    fraReadLine.Visible = bBool
    fraSkipLine.Visible = bBool

End Sub

Private Sub cmdCreateFile_Click()
    
    ' Create the file mentioned in the file name
    ' avoid any path if user has typed.
    
    Dim sFile As String
    Dim lRet As Long
    
    If txtCreatefile.Text = "" Then
        sFile = App.Path + "\" + "a.txt"
    Else    ' get the file name and append with the app.path
        sFile = App.Path + "\" + GetFileName(txtCreatefile.Text)
    End If
    
    'if the file already exist ask the user for overwrite
    If LCase(Dir(sFile, vbNormal)) = LCase(GetFileName(sFile)) Then
        lRet = MsgBox("File already exists. Do you want to overwrite?", vbOKCancel)
        If lRet = vbCancel Then Exit Sub
        
        ' if the file has readonly attribute remove that
        Set m_objFS = New Scripting.FileSystemObject
        Set m_objFile = m_objFS.GetFile(sFile)
        m_objFile.Attributes = Archive  ' 0
        
        Set m_objFile = Nothing
        Set m_objFS = Nothing
        DoEvents: DoEvents: DoEvents
    
    End If
           
    ' Create the file and write a comment line there
    Set m_objFS = New Scripting.FileSystemObject
    Set m_objTextStream = m_objFS.CreateTextFile(sFile, True)
    
    ' Write some information
    m_objTextStream.WriteLine Now
    m_objTextStream.WriteLine "This is the file created with FileSystem object "
    
    Set m_objFS = Nothing
    Set m_objTextStream = Nothing
    
End Sub

Private Sub ShowLayer(fraLayer As Frame)
    
On Error GoTo ErrShow
' Move the required layer to the left top and resize the form to its size
 
 fraLayer.Move 10, 10
 Me.Width = fraLayer.Left + fraLayer.Width + Con_lExtraWt
 Me.Height = fraLayer.Top + fraLayer.Height + Con_lExtraHt
 fraLayer.Visible = True
    
 Exit Sub
 
ErrShow:
    MsgBox "Error in Show Layer : " + Err.Description
End Sub

Private Sub cmdReadAll_Click()

On Error GoTo ErrRead

    ' Open the file and read all the contents in one go and display in the
    ' text box
    
    Dim sSampleFile As String
    
    sSampleFile = App.Path + "\" + ConSampleFile
    
    Set m_objFS = New Scripting.FileSystemObject
    Set m_objTextStream = m_objFS.OpenTextFile(sSampleFile, ForReading)
    
    txtReadAll.Text = ""
    txtReadAll.Text = m_objTextStream.ReadAll

    Set m_objFS = Nothing
    Set m_objTextStream = Nothing
    
    Exit Sub

ErrRead:
    MsgBox "CmdReadAll Click : " + Err.Description
End Sub

Private Sub cmdReadLine_Click()

On Error GoTo ErrRead

    Dim sSampleFile As String
    Dim sFromFile As String
    
    sSampleFile = App.Path + "\" + ConSampleFile

    Set m_objFS = New Scripting.FileSystemObject
    Set m_objTextStream = m_objFS.OpenTextFile(sSampleFile, ForReading)
    
    txtReadLine.Text = ""

    Do While Not m_objTextStream.AtEndOfStream
        ' Read the line from the text file
        sFromFile = m_objTextStream.ReadLine
        txtReadLine.Text = txtReadLine.Text + vbCrLf + sFromFile
    Loop

    Set m_objFS = Nothing
    Set m_objTextStream = Nothing
    
    Exit Sub
ErrRead:
    MsgBox "CmdReadLine click : " + Err.Description
End Sub

Private Sub cmdRead_Click()

On Error GoTo ErrRead

    Dim sSampleFile As String
    Dim iNumChars As Integer
    
    sSampleFile = App.Path + "\" + ConSampleFile
    
    iNumChars = txtReadChars.Text
    ' If nothing mentioned take 10 as the default
    If iNumChars = 0 Then iNumChars = 10
    
    Set m_objFS = New Scripting.FileSystemObject
    Set m_objTextStream = m_objFS.OpenTextFile(sSampleFile, ForReading)
    
    txtRead.Text = ""
    txtRead.Text = m_objTextStream.Read(iNumChars)
    
    Set m_objFS = Nothing
    Set m_objTextStream = Nothing

    Exit Sub

ErrRead:
    MsgBox "CmdRead Click : " + Err.Description
End Sub


Private Sub cmdSkipLine_Click()

On Error GoTo ErrSkip

    Dim sSampleFile As String
    Dim sFromFile As String
    Dim bSkip As Boolean
    
    sSampleFile = App.Path + "\" + ConSampleFile

    Set m_objFS = New Scripting.FileSystemObject
    Set m_objTextStream = m_objFS.OpenTextFile(sSampleFile, ForReading)

    txtSkipLine.Text = ""
    
    Do While Not m_objTextStream.AtEndOfStream
        ' Read the alternate lines from the text file
        If bSkip Then
            m_objTextStream.SkipLine
        Else
            sFromFile = m_objTextStream.ReadLine
            txtSkipLine.Text = txtSkipLine.Text + vbCrLf + sFromFile
        End If
        bSkip = Not bSkip
        
    Loop

    Set m_objFS = Nothing
    Set m_objTextStream = Nothing
    
    Exit Sub
ErrSkip:
    MsgBox "CmdSkipLine click : " + Err.Description
End Sub

Private Sub txtReadChars_KeyPress(KeyAscii As Integer)
    
On Error GoTo ErrReadChars

    ' Alow only  a 4 digit number. The sample file may not be so big.If more number is needed
    ' then change this code
    If Len(txtReadChars.Text) = 4 Then KeyAscii = 0: Exit Sub
    
    ' Allow only numeric and backspace here
    If KeyAscii = 8 Then
        ' do nothing
    ElseIf KeyAscii >= 48 And KeyAscii <= 57 Then
        ' do nothing
    Else    ' do not allow the character
        KeyAscii = 0
    End If
    
    Exit Sub
    
ErrReadChars:
    MsgBox "txtReadChars KeyPress : " + Err.Description
End Sub
