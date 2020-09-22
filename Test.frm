VERSION 5.00
Begin VB.Form fTest 
   Caption         =   "ucStatusbar 1.0 - Test"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   5820
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   251
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   388
   StartUpPosition =   2  'CenterScreen
   Begin Test.ucStatusbar ucStatusbar1 
      Height          =   630
      Left            =   150
      Top             =   3045
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1111
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   0
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuTestTop 
      Caption         =   "&Test"
      Begin VB.Menu mnuTest 
         Caption         =   "&Apply some changes"
         Index           =   0
      End
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    With ucStatusbar1
        
        '-- Initialize statusbar
        
        Call .Initialize(SizeGrip:=True, ToolTips:=True)
        
        '-- Initialize icons list
                             
        Call .InitializeIconList
        
        '-- Add icons
        
        Call .AddIcon(LoadResPicture("MAIL", vbResIcon))
        Call .AddIcon(LoadResPicture("USER", vbResIcon))
        Call .AddIcon(LoadResPicture("TIP", vbResIcon))
                             
        '-- Add panels
        
        Call .AddPanel(, , , [sbSpring], "Panel #1", , 0)
        Call .AddPanel(, 0, , [sbContents], "Panel #2", , 1)
        Call .AddPanel(, , , [sbSpring], "Last panel")
    End With
End Sub

Private Sub Form_Resize()
    
    ucStatusbar1.SizeGrip = (Me.WindowState <> vbMaximized)
End Sub

Private Sub mnuFile_Click(Index As Integer)

    Call Unload(Me)
End Sub

Private Sub mnuTest_Click(Index As Integer)

    With ucStatusbar1
        
        .Font.Size = 10 '* No effect when XP-theme enabled (control re-size)
        .Font.Bold = True
        
        .PanelStyle(1) = [sbPopOut]
        .PanelIconIndex(1) = 1
        
        .PanelIconIndex(2) = 0
        
        .PanelMinWidth(3) = 0
        .PanelAutosize(3) = [sbContents]
        .PanelText(3) = vbNullString
        .PanelTipText(3) = "Now you can see tool tip text"
        .PanelIconIndex(3) = 2
    End With
End Sub





Private Sub ucStatusbar1_PanelClick(ByVal Panel As Long, ByVal Button As MouseButtonConstants)
    Debug.Print "ucStatusbar1_PanelClick: Panel = " & Panel & " Button = " & Button
End Sub

Private Sub ucStatusbar1_PanelDblClick(ByVal Panel As Long, ByVal Button As MouseButtonConstants)
    Debug.Print "ucStatusbar1_PanelDblClick: Panel = " & Panel & " Button = " & Button
End Sub
