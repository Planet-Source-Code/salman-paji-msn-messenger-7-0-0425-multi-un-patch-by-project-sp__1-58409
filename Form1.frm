VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MSN 7.0.0425 Multi UN/Patch"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4455
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4455
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "UnPatch"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Patch"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Project SP : http://www.ptdnet.com/"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "MSN Messenger 7.0.0425 Multi UN/Patch by Project SP"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub MultiMSN_Patch()
Dim sp As Integer
sp = FreeFile
Open "msnmsgr.exe" For Binary As #sp
Put #sp, &H1181DC, "†"
Close #sp
End Sub
Public Sub MultiMSN_UnPatch()
Dim sp As Integer
sp = FreeFile
Open "msnmsgr.exe" For Binary As #sp
Put #sp, &H1181DC, "…"
Close #sp
End Sub
Private Sub Command1_Click()
Call MultiMSN_Patch
End Sub
Private Sub Command2_Click()
Call MultiMSN_UnPatch
End Sub
