VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "mini http client Sniffer"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   6330
      Width           =   8190
      _ExtentX        =   14446
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab st 
      Height          =   6255
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   11033
      _Version        =   393216
      TabOrientation  =   2
      TabHeight       =   1323
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Source"
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstHeaders"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txturl"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdGO"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtSource"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "UserAgent"
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "wb"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "|"
      TabPicture(2)   =   "frmMain.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2"
      Tab(2).Control(1)=   "Label1"
      Tab(2).ControlCount=   2
      Begin SHDocVwCtl.WebBrowser wb 
         Height          =   6135
         Left            =   -74160
         TabIndex        =   6
         Top             =   60
         Width           =   7215
         ExtentX         =   12726
         ExtentY         =   10821
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "res://C:\WINNT\system32\shdoclc.dll/dnserror.htm#http:///"
      End
      Begin VB.TextBox txtSource 
         Height          =   3555
         Left            =   840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2580
         Width           =   7155
      End
      Begin VB.CommandButton cmdGO 
         Caption         =   "GO"
         Default         =   -1  'True
         Height          =   315
         Left            =   7320
         TabIndex        =   4
         Top             =   120
         Width           =   675
      End
      Begin VB.TextBox txturl 
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Text            =   "http://"
         Top             =   120
         Width           =   6435
      End
      Begin VB.ListBox lstHeaders 
         Height          =   2010
         Left            =   840
         TabIndex        =   2
         Top             =   480
         Width           =   7155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "aprinzi@tiscalinet.it"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -72000
         MouseIcon       =   "frmMain.frx":1045
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   3240
         Width           =   1710
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Aldus the Prinzimaker                                Tiny HTTP CLIENT Sniffer - Beta 1.0   FreeWare"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -73020
         TabIndex        =   7
         Top             =   1920
         Width           =   5415
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------
' Ald Tiny HTTP Client Sniffer v.beta 1.0 FreeWare
'--------------------------------------------------------------------
' Author: Aldo Prinzi (Milan/Italy) e.mail:aprinzi@tiscalinet.it
' Date  : July 2001
'--------------------------------------------------------------------
'
' This Freeware version was written to demonstrate other things about
' the use of the Internet in a program written in Visual Basic.
'
' This is only a demonstration software freely distributed on an
' AS-IS basis. I can't know if it will work on any system or version
' of Visual Basic. It was written on VB 6.0 SP5 with Windows 2K Sp2.
'
' Needs Internet Explorer (4.x or later) on the system.
'--------------------------------------------------------------------
' I will be happy to share it with all who need this.
' Hope you like it.
' July 2001,
' Aldo Prinzi.
'--------------------------------------------------------------------
' "Microsoft", "Microsoft Windows" and "Microsoft Visual Basic"
' are trademark of Microsoft corp. - Redmond- USA
'--------------------------------------------------------------------
Option Explicit

Dim XHT As XMLHTTPRequest
Dim TmngA1 As Single
Dim TmngA2 As Single
Dim TmngB1 As Single
Dim TmngB2 As Single
Dim OnStart As Boolean
Dim OnNav As Boolean

Private Sub cmdGO_Click()
    Dim tmp As String
    Dim Pagina As String
    Dim I As Long
    Dim J As Long
    
    On Local Error Resume Next
    
    Set XHT = Nothing
    Set XHT = New XMLHTTPRequest
    
    lstHeaders.Clear
    txtSource = ""
    
    DoEvents
    TmngA1 = Timer
    If txturl <> "" Then
        XHT.open "GET", txturl, False
        XHT.send ""
        TmngA2 = Timer

        sb.Panels(1).Text = "Pure HTML doc dwnl time=" & Format(TmngA2 - TmngA1, "###0.##0") & " secs"
    
        Pagina = Replace(XHT.responseText, Chr(10), "")
        Pagina = Replace(Pagina, Chr(13), vbCrLf)
        txtSource = Pagina
        
        Pagina = XHT.getAllResponseHeaders
        J = 1
        Do
            I = InStr(J, Pagina, vbCrLf)
            If I = 0 Then Exit Do
            lstHeaders.AddItem Replace(Replace(Mid(Pagina, J, I - J + 2), Chr(13), ""), Chr(10), "")
            J = I + 1
        Loop
    Else
        st.Tab = 1
    End If
    If UCase(txturl) <> Left(Trim(UCase(wb.Document.URL)), Len(txturl)) Then
        wb.Navigate txturl
    End If

End Sub

Private Sub Form_Load()
    
    wb.Navigate "about:blank", False
    
    OnStart = True
    OnNav = False
    sb.Panels.Add
    sb.Panels(1).AutoSize = sbrSpring
    sb.Panels(2).AutoSize = sbrSpring
    st.TabCaption(2) = ""
    st.Tab = 0
    wb.Navigate "about:blank"
End Sub

Private Sub Form_Resize()
    Dim T As Integer
    T = st.Tab
    If Me.WindowState <> vbMinimized Then
        If Me.Width < 8000 Then Me.Width = 8000: Exit Sub
        If Me.Height < 7000 Then Me.Height = 7000: Exit Sub
        With st
            .Visible = False
            .Width = Me.ScaleWidth
            .Height = Me.ScaleHeight - sb.Height
            .Left = 0
            .Top = 0
        End With
        st.Tab = 1
        With wb
            .Width = st.Width - st.TabHeight - 60
            .Height = st.Height - 60
            .Top = 30
            .Left = st.TabHeight + 30
        End With
        st.Tab = 0
        lstHeaders.Width = st.Width - st.TabHeight - 60
        lstHeaders.Left = st.TabHeight + 30
        With txtSource
            .Width = st.Width - st.TabHeight - 60
            .Height = st.Height - 60 - .Top
            .Left = st.TabHeight + 30
        End With
        cmdGO.Left = st.Width + st.TabHeight - cmdGO.Width * 2 - 120
        txturl.Width = cmdGO.Left - txturl.Left - 60
        st.Tab = T
        st.Visible = True
    End If
End Sub

Private Sub Label2_Click()
    InputBox "Send me a mail using your own mailer. Simply copy the given address and paste it in the 'TO:' field of your mailer -> New Message", "Mail me", "aprinzi@tiscalinet.it"
End Sub

Private Sub wb_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    If OnNav = False Then
        TmngB1 = Timer
        OnNav = True
    End If
End Sub

Private Sub wb_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    TmngB2 = Timer
    If OnStart = True Then
        OnStart = False
    Else
        If UCase(txturl) <> Left(Trim(UCase(URL)), Len(txturl)) And LCase(Trim(URL)) <> "about:blank" Then
            txturl = URL
            cmdGO_Click
        End If
        sb.Panels(2).Text = "HTML browser time=" & Format(TmngB2 - TmngB1, "###0.##0") & " secs"
    End If
    OnNav = False
End Sub


