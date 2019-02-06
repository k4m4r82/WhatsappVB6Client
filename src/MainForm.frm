VERSION 5.00
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo WhatsApp Client for VB6"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5355
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtContact 
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox txtPesan 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   4935
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "Send"
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.OptionButton optUnknownContact 
      Caption         =   "Unknown Contact"
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin VB.OptionButton optContact 
      Caption         =   "Contact"
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton btnStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Contact/Phone Number"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Pesan"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2400
      Width           =   1695
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright (C) 2019 Kamarudin (http://coding4ever.net/)
'
' Licensed under the Apache License, Version 2.0 (the "License"); you may not
' use this file except in compliance with the License. You may obtain a copy of
' the License at
'
' http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
' WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the
' License for the specific language governing permissions and limitations under
' the License.
'
' The latest version of this file can be found at https://github.com/k4m4r82/WhatsappVB6Client
 
Option Explicit

Private client      As WhatsappVB6Client
Attribute client.VB_VarHelpID = -1

Private Sub btnSend_Click()
    Screen.MousePointer = vbHourglass
    DoEvents
    
    If optContact.Value Then
        Call client.SendToContact(txtContact.Text, txtPesan.Text)
    Else
        Call client.SendToUnknownContact(txtContact.Text, txtPesan.Text)
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub btnStart_Click()
    Dim url As String
    
    On Error GoTo errHandler
    
    Set client = New WhatsappVB6Client
    
    url = "https://web.whatsapp.com"
    
    Screen.MousePointer = vbHourglass
    DoEvents
    
    If client.Connect(url) Then
        Do While client.OnLoginPage
            DoEvents
        Loop
        
        btnStart.Enabled = False
        
        btnStop.Enabled = True
        btnSend.Enabled = True
    End If
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
errHandler:
    Screen.MousePointer = vbDefault
    Debug.Print Err.Description
End Sub

Private Sub btnStop_Click()
    Screen.MousePointer = vbHourglass
    DoEvents
    
    Call client.Disconnect
        
    btnStart.Enabled = True
        
    btnStop.Enabled = False
    btnSend.Enabled = False
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Terminate()
    If Not (client Is Nothing) Then
        Call client.Disconnect
        Call client.Dispose
    End If
End Sub
