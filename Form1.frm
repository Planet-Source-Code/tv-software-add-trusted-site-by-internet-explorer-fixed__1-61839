VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Trusted by Internet Explorer :) "
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "HKEY_LOCAL_MACHINE"
      Height          =   1215
      Left            =   0
      TabIndex        =   5
      Top             =   1800
      Width           =   4695
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   720
         TabIndex        =   7
         Top             =   360
         Width           =   3615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add"
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Site :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "HKEY_CURRENT_USER"
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   4695
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   720
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "Site :"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Add here your site to the trusted site's of Internet Explorer!"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
regCreate_A_Key HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\" & Text1.Text
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\" & Text1.Text, "*", 2
MsgBox "Your site is added to the trusted Site(s) of internet Explorer" & vbNewLine & "Don't forget to vote!", vbInformation, "Done"
End Sub

Private Sub Command2_Click()
regCreate_A_Key HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\" & Text2.Text
regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\" & Text2.Text, "*", 2
MsgBox "Your site is added to the trusted Site(s) of internet Explorer" & vbNewLine & "Don't forget to vote!", vbInformation, "Done"
End Sub

