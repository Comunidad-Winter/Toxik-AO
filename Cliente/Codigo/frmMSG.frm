VERSION 5.00
Begin VB.Form frmMSG 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Game Masters"
   ClientHeight    =   4905
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   2895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMsg 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2520
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   480
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3720
      Width           =   1995
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   180
      TabIndex        =   1
      Top             =   450
      Width           =   2520
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hint: hac� doble click en un usuario para ir hacia donde est� y borrar su mensaje autom�ticamente."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Han pedido ayuda..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1920
   End
   Begin VB.Menu menU_usuario 
      Caption         =   "Usuario"
      Visible         =   0   'False
      Begin VB.Menu mnuIR 
         Caption         =   "Ir donde esta el usuario"
      End
      Begin VB.Menu mnutraer 
         Caption         =   "Traer usuario"
      End
      Begin VB.Menu mnuBorrar 
         Caption         =   "Borrar mensaje"
      End
   End
End
Attribute VB_Name = "frmMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Dim lista As New Collection

Private Sub Form_Paint()
If Not cargado Then
    cargado = True
End If
End Sub

Public Sub MensajePoner(ByVal Nick As String, ByVal Mensaje As String)
On Error Resume Next
lista.Add Mensaje, Nick
End Sub

Public Sub MensajeBorrarTodos()
Do While lista.Count > 0
    Call lista.Remove(lista.Count)
Loop
End Sub

Private Sub Command1_Click()
Call MensajeBorrarTodos
Me.Visible = False
List1.Clear
End Sub

Private Sub Form_Load()
List1.Clear
AlwaysOnTop Me.hWnd
End Sub

Private Sub list1_Click()
On Error Resume Next
txtMsg.Text = lista.Item(List1.Text)
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    PopupMenu menU_usuario
End If

End Sub

Private Sub mnuBorrar_Click()
If List1.ListIndex < 0 Then Exit Sub
SendData ("SOSDON" & List1.List(List1.ListIndex))

List1.RemoveItem List1.ListIndex

End Sub

Private Sub mnuIR_Click()
SendData ("/IRA " & ReadField(1, List1.List(List1.ListIndex), Asc("-")))
End Sub

Private Sub mnutraer_Click()
SendData ("/SUM " & ReadField(1, List1.List(List1.ListIndex), Asc("-")))
End Sub

Private Sub list1_dblClick()
On Error Resume Next
SendData ("/IRA " & ReadField(1, List1.List(List1.ListIndex), Asc("-")))
SendData ("SOSDON" & List1.List(List1.ListIndex))
Call MensajeBorrarTodos
Me.Visible = False
List1.Clear
End Sub
