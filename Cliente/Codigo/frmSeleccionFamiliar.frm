VERSION 5.00
Begin VB.Form frmSeleccionFamiliar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Selecci�n de una mascota"
   ClientHeight    =   2745
   ClientLeft      =   3240
   ClientTop       =   3180
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picFamiliar 
      Height          =   1215
      Left            =   4830
      ScaleHeight     =   1155
      ScaleWidth      =   810
      TabIndex        =   7
      Top             =   210
      Width           =   870
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   4620
      TabIndex        =   4
      Top             =   2190
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4620
      TabIndex        =   3
      Top             =   1740
      Width           =   1335
   End
   Begin VB.TextBox txtFamiliarName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   4215
   End
   Begin VB.ComboBox lstFamiliar 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmSeleccionFamiliar.frx":0000
      Left            =   240
      List            =   "frmSeleccionFamiliar.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Label lblString 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de tu mascota"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label lblString 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de mascota"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSeleccionFamiliar.frx":0004
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmSeleccionFamiliar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmSeleccionFamiliar - ImperiumAO - v1.3.0
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

Private Sub Command1_Click()

CurrentUser.UserPet.Tipo = lstFamiliar.List(lstFamiliar.ListIndex)
CurrentUser.UserPet.nombre = txtFamiliarName.Text

If CurrentUser.UserPet.Tipo = "" Then
    MensajeAdvertencia ("�Selecciona tu familiar antes!")
    Exit Sub
ElseIf (CurrentUser.UserPet.nombre = "") Or Not AsciiValidos(CurrentUser.UserPet.nombre) Then
    MensajeAdvertencia ("Debe colocarle un nombre v�lido a su familiar")
    Exit Sub
End If

Call SendData("NFA" & CurrentUser.UserPet.Tipo & "," & CurrentUser.UserPet.nombre)

Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

ReDim ListaFamiliares(1 To 4) As tListaFamiliares
ListaFamiliares(1).Name = "Tigre"
ListaFamiliares(1).Desc = "Poseen grandes y filosas garras para atacar a tus oponentes."
ListaFamiliares(1).Imagen = "tigre.bmp"

ListaFamiliares(2).Name = "Lobo"
ListaFamiliares(2).Desc = "Astutos y arrogantes, su mordedura causa estragos en sus v�ctimas."
ListaFamiliares(2).Imagen = "lobo.bmp"

ListaFamiliares(3).Name = "Oso Pardo"
ListaFamiliares(3).Desc = "Se caracterizan por ser territoriales y muy resistentes."
ListaFamiliares(3).Imagen = "oso.bmp"

ListaFamiliares(4).Name = "Ent"
ListaFamiliares(4).Desc = "�Esta robusta criatura te defender� cual muro de piedra!"
ListaFamiliares(4).Imagen = "ent.bmp"

Dim i As Integer
lstFamiliar.Clear
lstFamiliar.AddItem ""
For i = LBound(ListaFamiliares) To UBound(ListaFamiliares)
    lstFamiliar.AddItem ListaFamiliares(i).Name
Next i

lstFamiliar.ListIndex = 0

End Sub

Private Sub lstFamiliar_Click()

If Not lstFamiliar.ListIndex = 0 Then
    picFamiliar.Picture = General_Load_Picture_From_Resource(ListaFamiliares(lstFamiliar.ListIndex).Imagen)
    Label1.Caption = ListaFamiliares(lstFamiliar.ListIndex).Desc
Else
    picFamiliar.Picture = Nothing
    Label1.Caption = "El tener 65 puntos en domar animales, te permite seleccionar una mascota que ser� acompa�ante de todas tus aventuras. Sea muy cuidadoso al seleccionar el tipo y el nombre, ya que �ste no podr� ser cambiado."
End If

End Sub
