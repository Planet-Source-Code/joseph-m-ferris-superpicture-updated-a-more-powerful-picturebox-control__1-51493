VERSION 5.00
Object = "{859DE455-7B89-457A-9743-C1081A14D235}#6.0#0"; "SuperPicture.ocx"
Begin VB.Form frmTest 
   Caption         =   "SuperPic Test Form"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin SuperPicture.SuperPicCtl SuperPicCtl1 
      Height          =   5070
      Left            =   240
      TabIndex        =   7
      Top             =   750
      Width           =   7710
      _extentx        =   13600
      _extenty        =   8943
   End
   Begin VB.ListBox lstEvents 
      Height          =   1230
      Left            =   0
      TabIndex        =   3
      Top             =   6015
      Width           =   8175
   End
   Begin VB.TextBox txtZoomLevel 
      Height          =   315
      Left            =   30
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   30
      Width           =   840
   End
   Begin VB.CommandButton cmdIn 
      Caption         =   "+"
      Height          =   285
      Left            =   945
      TabIndex        =   1
      Top             =   45
      Width           =   300
   End
   Begin VB.CommandButton cmdOut 
      Caption         =   "-"
      Height          =   285
      Left            =   1290
      TabIndex        =   0
      Top             =   45
      Width           =   300
   End
   Begin VB.CheckBox chkUseQuickBar 
      Caption         =   "Enable QuickBar"
      Height          =   210
      Left            =   3390
      TabIndex        =   5
      Top             =   105
      Value           =   1  'Checked
      Width           =   2010
   End
   Begin VB.CheckBox chkPanMode 
      Caption         =   "Panning Mode"
      Height          =   210
      Left            =   1785
      TabIndex        =   4
      Top             =   105
      Width           =   1425
   End
   Begin VB.Label lblHowTo 
      Caption         =   "Double-Click inside of the client area to load an image."
      Height          =   210
      Left            =   45
      TabIndex        =   6
      Top             =   390
      Width           =   7890
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*' Comments are sparse in here... Should be self-explanatory.  More detailed comments are in the control.
'*'
Option Explicit

Private Sub chkPanMode_Click()

    '*' Toggle pan mode based upon the check box.
    '*'
    If chkPanMode.Value = 1 Then
        SuperPicCtl1.PanActive = True
    Else
        SuperPicCtl1.PanActive = False
    End If
    
End Sub

Private Sub chkUseQuickBar_Click()

    '*' Toggle the use of the quickbar
    '*'
    If chkUseQuickBar.Value = 1 Then
        SuperPicCtl1.UseQuickBar = True
    Else
        SuperPicCtl1.UseQuickBar = False
    End If
    
End Sub

Private Sub cmdIn_Click()

    '*' Zoom in by an increment of 10% if the percentage is less than 100%
    '*'
    If SuperPicCtl1.Zoom < 1000 Then
        SuperPicCtl1.Zoom = SuperPicCtl1.Zoom + 10
    End If
    
End Sub

Private Sub cmdOut_Click()

    '*' Zoom out by an incrment of 10% if the percentage is more than 10%
    '*'
    If SuperPicCtl1.Zoom > 10 Then
        SuperPicCtl1.Zoom = SuperPicCtl1.Zoom - 10
    End If
    
End Sub

Private Sub Form_DblClick()

    '*' Unload the image.
    '*'
    SuperPicCtl1.UnloadImage
    
End Sub

Private Sub Form_Load()

    '*' By default, use the quickbar.
    '*'
    SuperPicCtl1.UseQuickBar = True
    
End Sub

Private Sub Form_Resize()

On Error Resume Next

    '*' Store the position if the window state is Max'd or Min'd.  Do it before resizing, since a restore will return
    '*' it to this size.
    '*'
    If Me.WindowState = vbMaximized Or Me.WindowState = vbMinimized Then
        SuperPicCtl1.StorePosition
    End If
    
    SuperPicCtl1.Move 0, 600, Me.ScaleWidth, Me.ScaleHeight - 660 - lstEvents.Height
    lstEvents.Move 0, Me.ScaleHeight - lstEvents.Height, Me.ScaleWidth, lstEvents.Height
    
    lstEvents.AddItem "Resize()"
    CleanList
    
    '*' Recall the position.  Will only work on Restore.
    '*'
    SuperPicCtl1.RecallPosition
        
End Sub

Private Sub SuperPicCtl1_Click()

    lstEvents.AddItem "Click()"
    CleanList
    
End Sub

Private Sub SuperPicCtl1_DblClick()

    lstEvents.AddItem "DblClick()"
    CleanList
    
    SuperPicCtl1.LoadImage
        
End Sub

Private Sub SuperPicCtl1_GotFocus()

    lstEvents.AddItem "GotFocus()"
    CleanList
    
End Sub

Private Sub SuperPicCtl1_KeyDown(KeyCode As Integer, Shift As Integer)

    lstEvents.AddItem "KeyDown(" & KeyCode & ", " & Shift & ")"
    CleanList
    
End Sub

Private Sub SuperPicCtl1_KeyPress(KeyAscii As Integer)

    lstEvents.AddItem "KeyPress(" & KeyAscii & ")"
    CleanList
    
End Sub

Private Sub SuperPicCtl1_KeyUp(KeyCode As Integer, Shift As Integer)

    lstEvents.AddItem "KeyUp(" & KeyCode & ", " & Shift & ")"
    CleanList
    
End Sub

Private Sub SuperPicCtl1_LostFocus()

    lstEvents.AddItem "LostFocus()"
    CleanList
    
End Sub

Private Sub SuperPicCtl1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lstEvents.AddItem "MouseDown(" & Button & ", " & Shift & ", " & X & ", " & Y & ")"
    CleanList
    
End Sub

Private Sub SuperPicCtl1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lstEvents.AddItem "MouseMove(" & Button & ", " & Shift & ", " & X & ", " & Y & ")"
    CleanList
    
End Sub

Private Sub SuperPicCtl1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lstEvents.AddItem "MouseUp(" & Button & ", " & Shift & ", " & X & ", " & Y & ")"
    CleanList
    
End Sub

Private Sub SuperPicCtl1_Paint()

    lstEvents.AddItem "Paint()"
    CleanList
    
End Sub

Private Sub SuperPicCtl1_Scroll()

    lstEvents.AddItem "Scroll()"
    CleanList
    
End Sub

Private Sub SuperPicCtl1_ZoomChanged(ByVal ZoomPercent As Long)

    lstEvents.AddItem "ZoomChanged(" & ZoomPercent & ")"
    CleanList
    
    '*' Toggle quickbar buttons based upon current percentage.
    '*'
    SuperPicCtl1.AllowZoomIn = (ZoomPercent < 1000)
    SuperPicCtl1.AllowZoomOut = (ZoomPercent > 10)
    
    txtZoomLevel.Text = ZoomPercent & "%"
    
End Sub

Private Sub SuperPicCtl1_ZoomInClick()

    lstEvents.AddItem "ZoomIn()"
    CleanList
    
    cmdIn_Click
    
End Sub

Private Sub SuperPicCtl1_ZoomOutClick()

    lstEvents.AddItem "ZoomOut()"
    CleanList
    
    cmdOut_Click
    
End Sub

Private Sub CleanList()

    If lstEvents.ListCount > 10 Then
        Do Until lstEvents.ListCount = 10
            Call lstEvents.RemoveItem(0)
        Loop
    End If
    
    lstEvents.ListIndex = lstEvents.ListCount - 1
    
End Sub
