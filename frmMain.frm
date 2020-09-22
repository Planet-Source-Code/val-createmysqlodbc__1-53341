VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Create MySql ODBC Connection Object"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7470
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Current Connections:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   26
      Top             =   5040
      Width           =   5175
      Begin VB.CommandButton btnDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4080
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox cmbDataSources 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   360
         Width           =   4935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MySql Server Information:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   20
      Top             =   2760
      Width           =   5175
      Begin VB.TextBox txtMySqlDBName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txtMySqlPassword 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   8
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox txtMySqlUserID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   7
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtMySqlPort 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   6
         Text            =   "3306"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtMySqlServer 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   5
         Text            =   "localhost"
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MySql Database Name:"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         Width           =   1680
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MySql Password:"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MySql User ID:"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MySql Port"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "My Sql Server:"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "ODBC Connection Information:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   6855
      Begin VB.TextBox txtMySqlStmt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   4
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtMySqlOption 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   3
         Text            =   "3"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtMySqlDriverName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   2
         Text            =   "C:\Windows\System32\myodbc3.dll"
         Top             =   1080
         Width           =   4335
      End
      Begin VB.TextBox txtODBCName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   0
         Top             =   360
         Width           =   4335
      End
      Begin VB.TextBox txtMySqlDescription 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   1
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MySql Stmt:"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MySql Option:"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MySql Driver Name:"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ODBC Connection Name:"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ODBC Connection Description:"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   2190
      End
   End
   Begin VB.CommandButton btnCreate 
      Caption         =   "&Create"
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Details about the MySql ODBC Connection you want to create"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   240
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   7290
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+--------------------------------------------------------------------+
'|                                                                    |
'|  Create a MySql ODBC Connection and delete any                     |
'|  ODBC Data Source programmatically                                 |
'|                                                                    |
'| This library is free software; you can redistribute it and/or      |
'| modify it at will                                                  |
'|                                                                    |
'| This library is distributed in the hope that it will be useful,    |
'| but WITHOUT ANY WARRANTY; without even the implied warranty of     |
'| MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.               |
'|                                                                    |
'| Created by Valere Palhories                                        |
'+--------------------------------------------------------------------+

Option Explicit

Private Sub btnCreate_Click()
    
    If CreateMySqlODBC(txtODBCName.Text, _
                        txtMySqlServer.Text, _
                        txtMySqlUserID.Text, _
                        txtMySqlPassword.Text, _
                        txtMySqlDBName.Text, _
                        txtMySqlDescription.Text, _
                        txtMySqlDriverName.Text, _
                        txtMySqlOption.Text, _
                        txtMySqlPort.Text, _
                        txtMySqlStmt.Text) = True Then
        GetDataSources cmbDataSources
        MsgBox txtODBCName.Text & " Create successfully", vbInformation
    End If
                    
End Sub

Private Sub btnDelete_Click()

    If cmbDataSources.ListIndex = -1 Then Exit Sub
    
    If MsgBox("Are you certain you want to delete " & _
                cmbDataSources.List(cmbDataSources.ListIndex) & "?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        DeleteMySqlODBC cmbDataSources.List(cmbDataSources.ListIndex)
        MsgBox cmbDataSources.List(cmbDataSources.ListIndex) & " deleted!", vbInformation
        cmbDataSources.RemoveItem (cmbDataSources.ListIndex)
    End If
    
End Sub

Private Sub Form_Load()

    GetDataSources cmbDataSources
    
End Sub
