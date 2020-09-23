VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Right-Click Example"
   ClientHeight    =   3015
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "Right-Click for information...                 Please Vote!!"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileCrap1 
         Caption         =   "This is what you see when you right click."
      End
      Begin VB.Menu mnuFileCrap2 
         Caption         =   "It's all very simple, just look at the code!"
      End
      Begin VB.Menu mnuFileCrap3 
         Caption         =   "All this text is coming from a premade menu"
      End
      Begin VB.Menu mnuFileCrap4 
         Caption         =   "called mnuFile."
      End
      Begin VB.Menu mnuFileCrap5 
         Caption         =   "If you goto the code and delete:"
      End
      Begin VB.Menu mnuFileCrap6 
         Caption         =   "           mnuFile.Visible = False"
      End
      Begin VB.Menu mnuFileCrap7 
         Caption         =   "under the FormLoad Sub, then you will"
      End
      Begin VB.Menu mnuFileCrap8 
         Caption         =   "be able to see the menu!"
      End
      Begin VB.Menu mnuFileCrap9 
         Caption         =   "PLEASE VOTE!!! Thanks!"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit Program"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
'This loads the form and makes mnuFile invisible..
'so all you see is when you right click...
mnuFile.Visible = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This sets the right click button..
If Button = 2 Then 'The right click button is button 2 [on most mouses]
    PopupMenu mnuFile 'Display mnuFile when you right click..
End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This is the same as above.. but this lets you see the menu
'when you right click the Label.. If this wasnt here then when you
'right clicked the menu, nothing would happen..
If Button = 2 Then
    PopupMenu mnuFile
End If
End Sub

Private Sub mnuFileExit_Click()
'Happens when you right click and click "Exit Program"
Unload Me
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Just exit crap
MsgBox "Thank you for trying my application.. I hope it proves usefull to you! Please vote. Thanks!", vbInformation, "Bye!"
Unload Me
End
End Sub
