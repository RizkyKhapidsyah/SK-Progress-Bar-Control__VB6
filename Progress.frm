VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Progress bar common control demo"
   ClientHeight    =   2835
   ClientLeft      =   2925
   ClientTop       =   3765
   ClientWidth     =   4590
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   189
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   306
   Begin VB.TextBox txtItins 
      Height          =   285
      Left            =   1140
      MaxLength       =   9
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   105
      Width           =   975
   End
   Begin VB.CheckBox chkVertical 
      Caption         =   "Vertical"
      Enabled         =   0   'False
      Height          =   225
      Left            =   1740
      TabIndex        =   3
      Top             =   510
      Width           =   855
   End
   Begin VB.CheckBox chkSmooth 
      Caption         =   "Smooth"
      Enabled         =   0   'False
      Height          =   225
      Left            =   780
      TabIndex        =   2
      Top             =   510
      Width           =   855
   End
   Begin VB.PictureBox picProgBarSize 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   660
      ScaleHeight     =   195
      ScaleWidth      =   3195
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2310
      Width           =   3195
   End
   Begin VB.CommandButton cmdDoStuff 
      Cancel          =   -1  'True
      Caption         =   "Do stuff that takes a while..."
      Default         =   -1  'True
      Height          =   345
      Left            =   180
      TabIndex        =   0
      Top             =   870
      Width           =   3195
   End
   Begin VB.Label labItins 
      Caption         =   "Itinerations:"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   150
      Width           =   855
   End
   Begin VB.Label labStyles 
      AutoSize        =   -1  'True
      Caption         =   "Styles:"
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   510
      Width           =   495
   End
   Begin VB.Label labPercent 
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   2310
      Width           =   375
   End
   Begin VB.Label labInfo 
      Caption         =   $"Progress.frx":0000
      Height          =   825
      Left            =   180
      TabIndex        =   6
      Top             =   1350
      Width           =   3375
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The progress bar's Form_Resize() mod level vars
Dim m_hProgBar As Long   ' hWnd
Dim m_dwPBLeft As Long   ' static x position when vertical
Dim m_dwPBTop As Long   ' static y position when horizontal
 

Private Sub Form_Load()
  
  ScaleMode = vbPixels
  cmdDoStuff.Cancel = True
  picProgBarSize.AutoRedraw = True   ' ...this is a demo...
  
  ' Restrict the txtItins input to digits only
  SetWindowLong txtItins.hwnd, GWL_STYLE, _
                          GetWindowLong(txtItins.hwnd, GWL_STYLE) Or _
                          ES_NUMBER
                          
  ' Set the initial itinerations
  txtItins = &H5000
  
  ' Make sure we can use the new IE3 styles
  chkSmooth.Enabled = IsNewComctl32()
  chkVertical.Enabled = chkSmooth.Enabled
  
End Sub

Private Sub cmdDoStuff_Click()
' Dynamically creates a progress bar and places it at the
' bottom (or the right) of the form, does stuff that takes a
' while & shows the progress, then destroys the progress bar.
' (the progress bar can be hidden w/ ShowWindow() instead
' of it being destroyed, but the time it takes to create it is
' negligible & it's resources are also freed w/ this method)

  Static bRunning As Boolean   ' cancel flag
  Dim bIsIE3 As Boolean
  Dim dwItins As Long
  Dim dwIncrement As Long
  Dim dwIdx As Long
  Dim vDummy As Variant
  On Error GoTo Out
  
  If bRunning Then bRunning = False: Exit Sub
      
  ' Create the progress bar using the standard window styles
  ' and the two new IE3 progress bar styles.
  ' Addition standard or extended window styles can be specified
  ' to alter the default appearance of the progress bar.
  ' The progress bar can also be easily created as a child window
  ' of a status bar part (VB "Panel"). Substitute the status bar's
  ' hWnd and a part's bounding rectangle (via SB_GETRECT) in
  ' CreateWindowEx()'s respective params below.
  m_hProgBar = CreateWindowEx(0, PROGRESS_CLASS, vbNullString, _
                                                  WS_CHILD Or WS_VISIBLE Or _
                                                  IIf(chkSmooth, PBS_SMOOTH, 0) Or _
                                                  IIf(chkVertical, PBS_VERTICAL, 0), _
                                                  0, 0, 0, 0, _
                                                  hwnd, 0, _
                                                  App.hInstance, ByVal 0)
  
  If m_hProgBar = 0 Then MsgBox "Uh oh...": Exit Sub
  
  ' Here we go...
  bRunning = True
  cmdDoStuff.Caption = "Stop"
  txtItins.Enabled = False
    
  ' Disable the IE3 style checkboxes during the progress
  ' (if Comctl32.dll's current working version is >= v4.70).
  ' The progress bar's styles can't be changed after it's been created.
  If chkSmooth.Enabled Then
    bIsIE3 = True   ' set the flag for re-enabling below
    chkSmooth.Enabled = False
    chkVertical.Enabled = False
  End If

  ' Set the progress bar's static x (or y) position
  ' so it's initially 15 pixels wide (or high)
  If chkVertical Then
    m_dwPBLeft = ScaleWidth - 15: m_dwPBTop = 0
  Else
    m_dwPBTop = ScaleHeight - 15: m_dwPBLeft = 0
  End If
  
  ' We'll do things a bit differently and let MoveWindow() in
  ' Form_Resize() set the progress bar's initial position & size.
  ' (no position & size values were specified in CreateWindowEx())
  ' Form_Resize() also displays the progress bar's current
  ' dimensions in picProgBarSize.
  Form_Resize
  
  dwItins = txtItins
  
  ' Set the range of the progess bar.
  ' (Minimum range = low word, Maximum range = high word).
  SendMessage m_hProgBar, PBM_SETRANGE, 0, ByVal (dwItins * &H10000)
  
  ' Set the value of the highlight increment. We''ll set it to 100
  ' itins here for the example even though it's the default value.
  dwIncrement = dwItins \ 100
  SendMessage m_hProgBar, PBM_SETSTEP, ByVal dwIncrement, 0
  
  ' Let's do some stuff...
  For dwIdx = 1 To dwItins
    DoEvents
    If Not bRunning Then Exit For
    
'    Sleep 1
    vDummy = vDummy & Format(Chr(Asc(Trim(Left("a", InStr(1, "a", "a", 1)))))) _
                                                                                             '\___/'
    If dwIdx Mod dwIncrement = 0 Then
      ' Advance the current position of the progress bar by the step increment.
      SendMessage m_hProgBar, PBM_STEPIT, 0, 0
      labPercent = dwIdx \ dwIncrement & "%"
    End If
    
    ' Either of these could be used instead of PBM_STEPIT above
    ' but the progress bar would be hit and redrawn on every itineration.
'    SendMessage m_hProgBar, PBM_SETPOS, ByVal dwIdx, 0
'    SendMessage m_hProgBar, PBM_DELTAPOS, ByVal 1, 0
    
  Next
  
Out:
  ' Frees all resources associated with the progress bar.
  ' If it's not destroyed here, the progress bar will automatically
  ' be destroyed when it's parent window (the window specified in
  ' the hWndParent param of CreateWindowEx()) is destroyed.
  If IsWindow(m_hProgBar) Then DestroyWindow m_hProgBar
  
  ' Re-initialize...
  bRunning = False
  cmdDoStuff.Caption = "Do stuff that takes a while..."
  txtItins.Enabled = True
  labPercent = ""
  picProgBarSize.Cls

  ' Eable the checkboxes if the flag was set above.
  chkSmooth.Enabled = bIsIE3
  chkVertical.Enabled = bIsIE3

End Sub

Private Sub Form_Resize()
  
  ' If we have a progress bar, adjust it's width & height w/ the form
  If IsWindow(m_hProgBar) Then
      MoveWindow m_hProgBar, m_dwPBLeft, _
                                              m_dwPBTop, _
                                              ScaleWidth - m_dwPBLeft, _
                                              ScaleHeight - m_dwPBTop, _
                                              True
    
    ' Display the progress bar's current size
    picProgBarSize.Cls
    picProgBarSize.Print "pixel width: " & ScaleWidth - m_dwPBLeft & _
                                 ", pixel height: " & ScaleHeight - m_dwPBTop
  Else
    picProgBarSize.Cls
    picProgBarSize.Print "(click the button to create a progress bar...)"
  End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set Form1 = Nothing
  End
End Sub
