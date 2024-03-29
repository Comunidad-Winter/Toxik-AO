VERSION 5.00
Begin VB.Form frmGoliath 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Operaci�n bancaria"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4590
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cerrar"
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Height          =   345
      Left            =   3000
      TabIndex        =   4
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox txtDatos 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   2460
      Width           =   4335
   End
   Begin VB.ListBox lstBanco 
      Height          =   840
      ItemData        =   "frmGoliath.frx":0000
      Left            =   90
      List            =   "frmGoliath.frx":0010
      TabIndex        =   1
      Top             =   1230
      Width           =   4395
   End
   Begin VB.Label lblDatos 
      Caption         =   "�Cu�nto deseas depositar?"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2130
      Width           =   4335
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGoliath.frx":0072
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4395
   End
End
Attribute VB_Name = "frmGoliath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmGoliath - ImperiumAO - v1.3.0
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation version 2.1 of
'the License
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'*****************************************************************

'*****************************************************************
'Augusto Jos� Rando (barrin@imperiumao.com.ar)
'   - First Relase
'*****************************************************************

Option Explicit

Private Oro As Long
Private Items As Long
Private CantTransferencia As Long
Private EtapaTransferencia As Byte

Public Sub ParseBancoInfo(ByVal rdata As String)

On Error GoTo Error_Handler

Oro = Val(General_Field_Read(1, rdata, ","))
Items = Val(General_Field_Read(2, rdata, ","))

If Val(Oro) > 0 And Val(Items) > 0 Then
    lblInfo.Caption = "Bienvenido a la cadena de finanzas Goliath. Tienes " & Items & " objetos en tu b�veda y " & Oro & " monedas de oro en tu cuenta... �C�mo te puedo ayudar?"
ElseIf Val(Oro) <= 0 And Val(Items) > 0 Then
    lblInfo.Caption = "Bienvenido a la cadena de finanzas Goliath. Tienes " & Items & " objetos en tu b�veda y a�n no has depositado oro... �C�mo te puedo ayudar?"
ElseIf Val(Oro) > 0 And Val(Items) <= 0 Then
    lblInfo.Caption = "Bienvenido a la cadena de finanzas Goliath. Tu b�veda est� vac�a y posees " & Oro & " monedas de oro en tu cuenta... �C�mo te puedo ayudar?"
ElseIf Val(Oro) <= 0 And Val(Items) <= 0 Then
    lblInfo.Caption = "Bienvenido a la cadena de finanzas Goliath. Tu b�veda y tu cuenta est�n vac�as... �C�mo te puedo ayudar?"
End If

Me.Show vbModeless, frmMain

Exit Sub

Error_Handler:
    'Error vite'

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()

Select Case lstBanco.ListIndex
    Case 0, -1 'Depositar
    
        'Negativos y ceros
        If (Val(txtDatos.Text) <= 0 And (UCase$(txtDatos.Text) <> "TODO")) Then lblDatos.Caption = "Cantidad inv�lida."
    
        If Val(txtDatos.Text) <= CurrentUser.UserGLD Or UCase$(txtDatos.Text) = "TODO" Then
            Call SendData("/DEPOSITAR " & IIf(Val(txtDatos.Text) > 0, Val(txtDatos.Text), CurrentUser.UserGLD))
            Unload Me
        Else
            lblDatos.Caption = "No tienes esa cantidad. Escr�bela nuevamente."
        End If
    Case 1 'Retirar
    
        'Negativos y ceros
        If (Val(txtDatos.Text) <= 0 And (UCase$(txtDatos.Text) <> "TODO")) Then lblDatos.Caption = "Cantidad inv�lida."
    
        If Val(txtDatos.Text) <= Oro Or UCase$(txtDatos.Text) = "TODO" Then
            Call SendData("/RETIRAR " & IIf(Val(txtDatos.Text) > 0, Val(txtDatos.Text), Oro))
            Unload Me
        Else
            lblDatos.Caption = "No tienes esa cantidad. Escr�bela nuevamente."
        End If
    Case 2 'B�veda
        Unload Me
    Case 3 'Transferir - Destino - Cantidad
        If EtapaTransferencia = 0 Then
        
            'Negativos y ceros
            If Val(txtDatos.Text) <= 0 Then
                lblDatos.Caption = "Cantidad inv�lida. Escr�bela nuevamente."
                txtDatos.Text = ""
                Exit Sub
            End If
            
            If Val(txtDatos.Text) <= Oro Then
                CantTransferencia = Val(txtDatos.Text)
                lblDatos.Caption = "�A qui�n le deseas enviar " & CantTransferencia & " monedas de oro?"
                EtapaTransferencia = 1
                txtDatos.Text = ""
            Else
                lblDatos.Caption = "No tienes esa cantidad. Escr�bela nuevamente."
                txtDatos.Text = ""
            End If
        ElseIf EtapaTransferencia = 1 Then
            If txtDatos.Text <> "" Then
                Call SendData("TRA" & txtDatos.Text & ";" & CantTransferencia)
                Unload Me
            Else
                lblDatos.Caption = "�Nombre de destino inv�lido!"
                txtDatos.Text = ""
            End If
        End If
End Select

End Sub

Private Sub lstBanco_Click()

Select Case lstBanco.ListIndex
    Case 0 'Depositar
        lblDatos.Caption = "�Cu�nto deseas depositar?"
    Case 1 'Retirar
        lblDatos.Caption = "�Cu�nto deseas retirar?"
    Case 2 'B�veda
        Call SendData("INITBOV")
        Unload Me
    Case 3 'Transferir
        EtapaTransferencia = 0
        lblDatos.Caption = "�Qu� cantidad deseas transferir?"
End Select

End Sub
