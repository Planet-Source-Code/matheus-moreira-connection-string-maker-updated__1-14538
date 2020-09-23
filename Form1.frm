VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADO Connection Class Demonstration Project"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Database Properties"
      Height          =   2295
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   5055
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   600
         Width           =   3495
      End
      Begin VB.CheckBox chkUseNTSecurity 
         Caption         =   "Use NT Security to Login"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Username:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Database Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "ServerName:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   930
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   2520
      Width           =   7215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Database Type"
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
      Begin VB.OptionButton optProvider 
         Caption         =   "Excel 8.0 Files (XLS)"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Width           =   1815
      End
      Begin VB.OptionButton optProvider 
         Caption         =   "Dbase III Files (DBF)"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   1815
      End
      Begin VB.OptionButton optProvider 
         Caption         =   "SQL Server"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optProvider 
         Caption         =   "Access 2000 (Jet 4.0)"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1875
      End
      Begin VB.OptionButton optProvider 
         Caption         =   "Access 95 (Jet 3.51)"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Tag             =   "0"
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Conn As New AdoConnectionClass

Private Sub chkUseNTSecurity_Click()
If chkUseNTSecurity.Value = 0 Then
    EnableControl "label4"
    EnableControl "Text4"

    EnableControl "label5"
    EnableControl "Text5"
Else
    DisableControl "label4"
    DisableControl "Text4"

    DisableControl "label5"
    DisableControl "Text5"
End If
End Sub
Private Sub Command1_Click()
On Error Resume Next
Dim PopGrid As ADODB.Recordset
Dim RandonAccessSQLServerConnection As Integer

Text1 = ""

RandonAccessSQLServerConnection = Rnd * 1

Select Case optProvider(0).Tag
    Case 0, 1 'Open an Access 95-2000 database
        'The lines below do the same thing.
        'the OpenAccess method it's only a shortcut.
        If RandonAccessSQLServerConnection = 1 Then
            'Function OpenAccess open only JET 4.0
            If Not Conn.OpenAccess(Text3.Text, Text5.Text) Then GoTo Failed
        Else
            If optProvider(0).Tag = 0 Then
                Conn.ProviderConst = pdsajet
            Else
                Conn.ProviderConst = pdsajet40
            End If
            Conn.DataSource = Text3.Text
            Conn.Password = Text5.Text
            If Not Conn.DataOpen Then GoTo Failed
        End If
        
    Case 2 'Open an SQL Server database
        'The lines below do the same thing.
        'the OpenSQLServer method it's only a shortcut.
        If RandonAccessSQLServerConnection = 1 Then
            If Not Conn.OpenSQLServer(Text2.Text, Text3.Text, Text4.Text, Text5.Text, IIf(chkUseNTSecurity.Value = 0, False, True)) Then GoTo Failed
        Else
            Conn.ProviderConst = pdsasqlserver
            Conn.DataSource = Text2.Text
            Conn.InitialCatalog = Text3.Text
            Conn.UseNTSecurity = IIf(chkUseNTSecurity.Value = 0, False, True)
            Conn.UserID = Text4.Text
            Conn.Password = Text5.Text
            If Not Conn.DataOpen Then GoTo Failed
        End If

    Case 3 'Open an Dbase III database directory
        Conn.ProviderConst = pdsadbase
        Conn.DataSource = Text3.Text
        If Not Conn.DataOpen Then GoTo Failed
    
    Case 4 'Open an Excel 8.0 database directory
        Conn.ProviderConst = pdsaexcel
        Conn.DataSource = Text3.Text
        If Not Conn.DataOpen Then GoTo Failed
    
End Select

Set PopGrid = Conn.Connection.OpenSchema(adSchemaTables)
Text1 = "List of Tables on Database:" & vbCrLf & "---------------------------" & vbCrLf

Do While Not PopGrid.EOF
    Text1.Text = Text1.Text & PopGrid.Fields("TABLE_NAME").Value & vbCrLf
    PopGrid.MoveNext
Loop

Exit Sub
Failed:
MsgBox "The database could not be opened!", vbCritical, "Error"

End Sub

Private Sub Form_Load()
chkUseNTSecurity.Enabled = False
DisableControl "label2"
DisableControl "Text2"
DisableControl "label3"
DisableControl "Text3"
DisableControl "label4"
DisableControl "Text4"
DisableControl "label5"
DisableControl "Text5"
End Sub

Private Sub optProvider_Click(Index As Integer)
optProvider(0).Tag = Index

chkUseNTSecurity.Value = 0
Select Case Index
    Case 0, 1 'Access 95 - 2000
        Label3 = "Database Name:"
        
        If Index = 0 Then
            Conn.ProviderConst = pdsajet
        Else
            Conn.ProviderConst = pdsajet40
        End If
        
        chkUseNTSecurity.Enabled = False
        
        DisableControl "label2"
        DisableControl "Text2"
        
        EnableControl "label3"
        EnableControl "Text3"
        
        DisableControl "label4"
        DisableControl "Text4"
        
        EnableControl "label5"
        EnableControl "Text5"
        
    Case 2 'SQL Server
        Label3 = "Database Name:"
        
        Conn.ProviderConst = pdsasqlserver
        
        chkUseNTSecurity.Enabled = True
        
        EnableControl "label2"
        EnableControl "Text2"
        
        EnableControl "label3"
        EnableControl "Text3"
        
        EnableControl "label4"
        EnableControl "Text4"
        
        EnableControl "label5"
        EnableControl "Text5"
    
    Case 3, 4
        If Index = 3 Then
            Label3 = "DBFs Directory:"
            Conn.ProviderConst = pdsadbase
        Else
            Label3 = "XLS Name:"
            Conn.ProviderConst = pdsaexcel
        End If
        chkUseNTSecurity.Enabled = False
        
        DisableControl "label2"
        DisableControl "Text2"
        
        EnableControl "label3"
        EnableControl "Text3"
        
        DisableControl "label4"
        DisableControl "Text4"
        
        DisableControl "label5"
        DisableControl "Text5"
        
End Select

End Sub
Private Sub DisableControl(ByVal ControlName As String)
    If TypeOf Me.Controls(ControlName) Is VB.Label Then
        Me.Controls(ControlName).ForeColor = vbGrayText
    Else
        Me.Controls(ControlName).ForeColor = vbGrayText
        Me.Controls(ControlName).BackColor = vbButtonFace
    End If
    Me.Controls(ControlName).Enabled = False
End Sub
Private Sub EnableControl(ByVal ControlName As String)
    If TypeOf Me.Controls(ControlName) Is VB.Label Then
        Me.Controls(ControlName).ForeColor = vbWindowText
    Else
        Me.Controls(ControlName).ForeColor = vbWindowText
        Me.Controls(ControlName).BackColor = vbWindowBackground
    End If
    Me.Controls(ControlName).Enabled = True
End Sub

