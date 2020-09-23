VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Ping/Resolve Example"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Hostname to IP"
      Height          =   3135
      Left            =   2760
      TabIndex        =   6
      Top             =   240
      Width           =   2535
      Begin VB.CommandButton cmdPingHostname 
         Caption         =   "Ping Hostname"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   2175
      End
      Begin VB.ListBox lstResolvedAddress 
         Height          =   450
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txtHostname 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Text            =   "www.microsoft.com"
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton cmdResolveIp 
         Caption         =   "Resolve IP Addresses"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label lblHostPingState 
         AutoSize        =   -1  'True
         Caption         =   "Ping State:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   2760
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Enter Hostname"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Resolved IP Addresses"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1650
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "IP to Hostname"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2535
      Begin VB.CommandButton cmdPingIp 
         Caption         =   "Ping IP Address"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   2280
         Width           =   2175
      End
      Begin VB.CommandButton cmbResolveHostname 
         Caption         =   "Resolve Hostname"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox txtResolvedHostname 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txtIpAddress 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Text            =   "207.46.131.137"
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblIpPingState 
         AutoSize        =   -1  'True
         Caption         =   "Ping State:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   2760
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Resolved Hostname"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Enter IP Address"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'IP to Hostname Section
Private Sub cmbResolveHostname_Click()
    txtResolvedHostname = ResolveHostname(txtIpAddress)
End Sub
Private Sub cmdPingIp_Click()
    lblIpPingState.Caption = "Ping State: " & IIf(Ping(txtIpAddress, 1000), "Alive!", "Dead!")
End Sub

'Hostname to IP Section
Private Sub cmdResolveIp_Click()
Dim retColl As Collection
Dim nCount As Integer

    Set retColl = ResolveIpaddress(txtHostname)
    
    lstResolvedAddress.Clear
    If retColl.Count > 0 Then
        For nCount = 1 To retColl.Count
            lstResolvedAddress.AddItem CStr(retColl.Item(nCount))
        Next nCount
    End If

End Sub
Private Sub cmdPingHostname_Click()
    lblHostPingState.Caption = "Ping State: " & IIf(Ping(txtHostname, 1000), "Alive!", "Dead!")
End Sub


