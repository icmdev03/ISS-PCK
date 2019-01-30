VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm frmMDI 
   AutoShowChildren=   0   'False
   BackColor       =   &H80000004&
   Caption         =   "iSmartSales Order Interface"
   ClientHeight    =   5985
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12525
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12525
      _ExtentX        =   22093
      _ExtentY        =   1588
      ButtonWidth     =   1720
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   2
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "‚À≈¥ÕÕ‡¥Õ√Ï"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "ÕÕ°"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   1680
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMain.frx":55266C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu1 
      Caption         =   "‚À≈¥"
      Begin VB.Menu mnu11 
         Caption         =   "‚À≈¥ÕÕ‡¥Õ√Ï"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()

End Sub

'Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    ShutDownApplication
'End Sub

Private Sub mnu11_Click()
'    frmLoadOrder.Show
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)


    Select Case Button

        Case "‚À≈¥ÕÕ‡¥Õ√Ï"
            Call Main
'            gstrUserType = ""
'            frmLoadOrder.Show

        Case "ÕÕ°"
            End
'            ShutDownApplication
    End Select
End Sub
