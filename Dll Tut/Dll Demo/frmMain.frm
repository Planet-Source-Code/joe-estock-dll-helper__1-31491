VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Dll Demo"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtWindowsDirectory 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   2895
   End
   Begin VB.CommandButton cmdWindowsDirectory 
      Caption         =   "Get Windows Directory"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox txtSystemDirectory 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2895
   End
   Begin VB.CommandButton cmdSystemDirectory 
      Caption         =   "Get System Directory"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This is irrelevant to the program, however I wanted to state
'something (as I do with every code I release or submit).
'I am not going to beg you to vote for me, and I never will.
'I feel that if you think this code is of any use to you,
'then you will vote for it accordingly. However, voting for me
'provides a way for me to know how useful a particular
'code is and by that I know what to work on next. I also highly
'appreciate your comments, suggestions, as well as complaints
'and will answer them to the best of my knowledge. Feel free to
'contact me via email at joeestock@hotmail.com, however I prefer
'you to post your feedback on the Planet Source Code site so that
'others may benefit from the questions you may ask, or the comments
'you may post. This greatly helps the community as well as developers
'alike. Thank you for your support and patronage, and as always I
'promise to deliver valuable and informational code.
'
'This code is NOT copyrighted in any way and I urge you to
'use it in whatever project you may need it in. I do not even
'ask that you add me into the credits, or even source code
'of your program.
'
'Random Thought: Why does frozen water make so much money?
'
'       Joe Estock

'First you must register the dynamic link library (dll). To do this,
'Go to Start, Run, Then type in (with out the brackets [])
'[regsvr32 path and filename of dll to register]
'then type enter. Now add a reference to the dll
'by selecting Project, References withing visual basic.
'Click browse, then browse to the path of the dll you
'just registered.

Private Sub cmdSystemDirectory_Click()
    'Declare a variable to reference our new dll
    Dim sDllTut As New DirectoryTools.clsDirectoryTools
    
    'Set the text of the textbox to the value
    'returned by the GetSysDir function from
    'within the dll
    txtSystemDirectory.Text = sDllTut.GetSysDir
    
    'Unreference the variable so that we do not
    'hold up memory for something we no longer need
    Set sDllTut = Nothing
    
    'Make for a gracefull exit, grab a jolt cola, and
    'make another program with your newly obtained knowledge
End Sub

Private Sub cmdWindowsDirectory_Click()
    'Declare a variable to reference our new dll
    Dim sDllTut As New DirectoryTools.clsDirectoryTools
    
    'Set the text of the textbox to the value
    'returned by the GetSysDir function from
    'within the dll
    txtWindowsDirectory.Text = sDllTut.GetWinDir
    
    'Unreference the variable so that we do not
    'hold up memory for something we no longer need
    Set sDllTut = Nothing
    
    'Make for a gracefull exit, grab a jolt cola, and
    'make another program with your newly obtained knowledge
End Sub
