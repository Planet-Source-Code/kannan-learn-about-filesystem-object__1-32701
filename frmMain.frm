VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5790
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   5790
   ScaleWidth      =   8655
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTopicDescription 
      BorderStyle     =   0  'None
      Height          =   2745
      Left            =   3855
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   1590
      Width           =   3465
   End
   Begin VB.Image imgViewSample 
      Height          =   450
      Left            =   5685
      Picture         =   "frmMain.frx":145FF
      Top             =   4515
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label lblTopicTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4470
      TabIndex        =   5
      Top             =   1110
      Width           =   2280
   End
   Begin VB.Line lineVisual 
      BorderColor     =   &H00C0FFC0&
      Visible         =   0   'False
      X1              =   2745
      X2              =   3825
      Y1              =   1245
      Y2              =   1245
   End
   Begin VB.Line lineCorner 
      BorderColor     =   &H0080FF80&
      Index           =   3
      X1              =   6780
      X2              =   7320
      Y1              =   4380
      Y2              =   4380
   End
   Begin VB.Line lineCorner 
      BorderColor     =   &H0080FF80&
      Index           =   2
      X1              =   3840
      X2              =   4380
      Y1              =   4365
      Y2              =   4365
   End
   Begin VB.Line lineCorner 
      BorderColor     =   &H0080FF80&
      Index           =   0
      X1              =   6780
      X2              =   7320
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line lineCorner 
      BorderColor     =   &H0080FF80&
      Index           =   1
      X1              =   3825
      X2              =   4365
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FF80&
      Index           =   2
      X1              =   7320
      X2              =   7320
      Y1              =   1200
      Y2              =   4380
   End
   Begin VB.Line lineSide 
      BorderColor     =   &H0080FF80&
      Index           =   1
      X1              =   3825
      X2              =   3825
      Y1              =   1200
      Y2              =   4380
   End
   Begin VB.Image imgExit 
      Height          =   450
      Left            =   7305
      Picture         =   "frmMain.frx":152B9
      Top             =   5205
      Width           =   1200
   End
   Begin VB.Image imgExitDown 
      Height          =   450
      Left            =   405
      Picture         =   "frmMain.frx":15F21
      Top             =   6750
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgExitNormal 
      Height          =   450
      Left            =   405
      Picture         =   "frmMain.frx":16B89
      Top             =   6105
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label lblRight 
      Height          =   4575
      Left            =   7920
      TabIndex        =   4
      Top             =   660
      Width           =   735
   End
   Begin VB.Label lblLeft 
      Height          =   4680
      Left            =   0
      TabIndex        =   3
      Top             =   645
      Width           =   735
   End
   Begin VB.Label lblBottom 
      Height          =   690
      Left            =   0
      TabIndex        =   2
      Top             =   5115
      Width           =   8670
   End
   Begin VB.Label lblTop 
      Height          =   750
      Left            =   15
      TabIndex        =   1
      Top             =   -15
      Width           =   8685
   End
   Begin VB.Label lblTopic 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Topic1"
      ForeColor       =   &H000080FF&
      Height          =   195
      Index           =   0
      Left            =   1305
      TabIndex        =   0
      Top             =   1020
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


' To move the window as you like

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
' **********

' Just to have a delay
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private Const m_Capas = "Label1"
Private m_iCurTopic As Integer
Private m_bTopicSelected As Boolean

Private Sub Form_Load()
    
' Do the Initial settings here

    ' Set the file system object here
    Set g_objFileSystem = New Scripting.FileSystemObject
    ' g_objTextStream  object will be used when we open a text file
    
    ' Set the border labels transparent
    lblLeft.BackStyle = 0
    lblRight.BackStyle = 0
    lblTop.BackStyle = 0
    lblBottom.BackStyle = 0
    
    ' Set the default picture to exit
    imgExit.Picture = imgExitNormal.Picture
    
    If LoadTopics Then
        lblTopic_Click 1
    Else
        MsgBox "Cannot read and update the data"
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
'''''    If m_bTopicSelected Then
'''''        '
'''''
'''''    Else        ' Reset the topic info
'''''
'''''        ' Make the highlighted topics (if any) to normal
'''''
'''''        m_iCurTopic = 0
'''''
'''''        MakeLabelBGNormal
'''''
'''''        ' Make the exit button picture to normal
'''''        ImgExitNormalPicture
'''''
'''''        ' Hide the line connecting the lable and the topic description area
'''''        lineVisual.Visible = False
'''''    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Clear the array and the objects used
    ReDim g_arrTopics(0)
    
    Set g_objFileSystem = Nothing
    
End Sub

Private Sub imgExit_Click()
    Unload Me
End Sub

Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        ImgExitNormalPicture
        
'        If imgExit.Picture <> imgExitNormal.Picture Then
'            imgExit.Picture = imgExitNormal.Picture
'        End If
    End If

End Sub


Private Sub imgExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbLeftButton Then
        ImgExitNormalPicture
    Else
        ImgExitDownPicture
    End If
    
End Sub


Private Sub imgExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
'        ImgExitNormalPicture
        
'        If imgExit.Picture <> imgExitNormal.Picture Then
'            imgExit.Picture = imgExitNormal.Picture
'        End If
    End If

End Sub


Private Sub imgViewSample_Click()

    Debug.Print m_iCurTopic

    Load frmSample
    
'''    Select Case m_iCurTopic
'''
'''        Case 2      ' CreateFile
'''
'''        Case 4      'Read
'''
'''        Case 5      'ReadAll
'''
'''        Case 6      ' ReadLine
'''
'''        Case 8      ' SkipLine
'''
'''    End Select

    frmSample.CurrentTopic = m_iCurTopic
    Me.Hide
    frmSample.Show 1
    Me.Visible = True

End Sub

Private Sub lblBottom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ImgExitNormalPicture
  If Button = 1 Then
    DragForm
  End If
  
End Sub


Private Sub lblLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ImgExitNormalPicture
  If Button = 1 Then
    DragForm
  End If

End Sub


Private Sub lblRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ImgExitNormalPicture
  If Button = 1 Then
    DragForm
  End If
End Sub


Private Sub DragForm()

Dim lngReturnValue As Long
    
    Call ReleaseCapture
    lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)

End Sub

Private Sub lblTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ImgExitNormalPicture
  If Button = 1 Then
    DragForm
  End If
  
End Sub

Private Sub lblTopic_Click(Index As Integer)

    m_iCurTopic = Index
    m_bTopicSelected = True

    MakeLabelBGNormal

     ' Highlight the current topic
    lblTopic(Index).BackColor = &HC0FFC0


    lineVisual.Visible = False

    ' Move the visual connector line to the current location
    lineVisual.X1 = lblTopic(Index).Left + lblTopic(Index).Width
    lineVisual.Y1 = lblTopic(Index).Top + (lblTopic(Index).Height / 2)

    lineVisual.X2 = lineSide(1).X1
    lineVisual.Y2 = lineVisual.Y1
    lineVisual.Visible = True

    DisplayTopic

End Sub

Private Sub lblTopic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
     
     ' If the last selected topic is something else
     If m_bTopicSelected And m_iCurTopic <> Index Then m_bTopicSelected = False
     
     ' set the current label
     'm_iCurTopic = Index
     
'''''     ' Make other topics bg as normal
'''''     MakeLabelBGNormal
'''''
'''''     ' Highlight the current topic
'''''     lblTopic(Index).BackColor = &HC0FFC0
'''''
'''''     lineVisual.Visible = False
'''''
'''''     ' Move the visual connector line to the current location
'''''     lineVisual.X1 = lblTopic(Index).Left + lblTopic(Index).Width
'''''     lineVisual.Y1 = lblTopic(Index).Top + (lblTopic(Index).Height / 2)
'''''
'''''     lineVisual.X2 = lineSide(1).X1
'''''     lineVisual.Y2 = lineVisual.Y1
'''''     lineVisual.Visible = True
'''''
'''''     DisplayTopic

End Sub

Private Sub MakeLabelBGNormal()
Dim iTemp As Integer

    If g__iNumTopics > 0 Then
        For iTemp = 1 To g__iNumTopics
            If m_iCurTopic <> iTemp Then
                lblTopic(iTemp).BackColor = &H80000009
            End If
        Next
    End If
    
End Sub

Private Sub ImgExitNormalPicture()

    If imgExit.Picture <> imgExitNormal.Picture Then
        imgExit.Picture = imgExitNormal.Picture
    End If

End Sub

Private Sub ImgExitDownPicture()

    If imgExit.Picture <> imgExitDown.Picture Then
        imgExit.Picture = imgExitDown.Picture
    End If

End Sub


Private Function LoadTopics() As Boolean

' Read the topic from the file , fill the data in the array and create a
' new interface to access the topic. Here the interface is label.

On Error GoTo errLoad

Dim iTemp As Integer
Dim iCount As Integer
Dim sTemp As String


    If Not ReadAndFillArray Then Exit Function
    
    iCount = g__iNumTopics
    
    
    For iTemp = 1 To iCount  ' topic 1 label is already loaded
        
        Load lblTopic(iTemp)
        lblTopic(iTemp).Move lblTopic(iTemp - 1).Left, lblTopic(iTemp - 1).Top + lblTopic(iTemp - 1).Height + 20
        
        sTemp = g_arrTopics(iTemp).sTopicTitle
        If sTemp = "" Then sTemp = "     " ' Just to make the topic label visible to some extent
        
        lblTopic(iTemp).Caption = sTemp
        lblTopic(iTemp).Visible = True
    Next
    
    LoadTopics = True
    
    Exit Function
    
errLoad:

    Debug.Print Err.Description
    Resume Next
    
End Function


Private Sub DisplayTopic()
    
    lblTopicTitle.Caption = g_arrTopics(m_iCurTopic).sTopicTitle
    txtTopicDescription.Text = g_arrTopics(m_iCurTopic).sTopicDescription
    
    imgViewSample.Visible = IIf(g_arrTopics(m_iCurTopic).bCodeSampleAvailable, True, False)

End Sub
