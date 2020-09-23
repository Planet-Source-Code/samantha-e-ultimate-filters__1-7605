VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmWaiting 
   ClientHeight    =   825
   ClientLeft      =   4515
   ClientTop       =   4590
   ClientWidth     =   5250
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   5250
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Applying filter. Please, wait."
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "frmWaiting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

