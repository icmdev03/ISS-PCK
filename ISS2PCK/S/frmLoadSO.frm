VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmLoadSO 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "iSmartSales Quotation ,Order Interface V1.4"
   ClientHeight    =   2805
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8625
   Icon            =   "frmLoadSO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   8625
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   1588
      ButtonWidth     =   1720
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   3
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "โหลดใบเสนอ"
            Key             =   "Loadquote"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "โหลดออเดอร์"
            Key             =   "Loadso"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "ออก"
            Key             =   "Exit"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Lblstatus 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ก่อนรันโปรแกรมให้แจ้งทุกท่าน ออกจาก Express ครับ !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   8415
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLoadSO.frx":0442
            Key             =   "Loadso"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLoadSO.frx":55266C
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLoadSO.frx":552BBE
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmLoadSO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Terminate()

    Unload Me
    End
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)

    Dim Response

    Select Case Button.Index
        Case 1
'            If (Format(Time, "hh:mm") > "22:30" And Format(Time, "hh:mm") < "10:15") Then
'
'                MsgBox "ไม่อนุญาติให้รันช่วงเวลา 22:30-10:15  เนื่องจากรัน Interface ครับ !", vbOKOnly
'            Else
            Response = MsgBox("ท่านต้องการ Run โปรแกรมใบเสนอราคา ใช่ หรือ ไม่?", vbYesNo, "Do you want to run Program?")
            If Response = vbYes Then
                Lblstatus.Caption = "กำลัง Run โปรแกรมอยู่  .  .  .  ."
                frmLoadSO.Refresh
                strThaidoc = "ใบเสนอราคา"
                strEngdoc = "Quotation"
                strRuntype = "QT"
                Call Main
                Lblstatus.Caption = " โปรแกรม Run เสร็จเรียบร้อย"
            End If
'            End If

        Case 2
            Response = MsgBox("ท่านต้องการ Run โปรแกรมใบสั่งขาย ใช่ หรือ ไม่?", vbYesNo, "Do you want to run Program?")
            If Response = vbYes Then
                Lblstatus.Caption = "กำลัง Run โปรแกรมอยู่  .  .  .  ."
                frmLoadSO.Refresh
                strThaidoc = "ใบสั่งขาย"
                strEngdoc = "Sales Note"
                strRuntype = "SO"
                Call Main
                Lblstatus.Caption = " โปรแกรม Run เสร็จเรียบร้อย"
            End If
            
        Case 3
            Unload Me
            End
    End Select
    
End Sub
