VERSION 5.00
Begin VB.Form frmMapa 
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMapa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8715
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11730
      Begin VB.PictureBox picMapa 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   13000
         Left            =   0
         ScaleHeight     =   867
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1034
         TabIndex        =   3
         Top             =   0
         Width           =   15510
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   15
      TabIndex        =   1
      Top             =   8745
      Width           =   11715
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   8745
      Left            =   11730
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmMapa - ImperiumAO - v1.3.0
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

Private Sub picMapa_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If (Button = vbLeftButton) Then
    Call Auto_Drag(Me.hwnd)
Else
    Unload Me
End If

End Sub

Private Sub Form_Load()

    picMapa.Picture = General_Load_Picture_From_Resource("mapa.bmp")
    Call Make_Transparent_Form(Me.hwnd, 210)

    With HScroll1
        .min = 1
        .max = picMapa.Width - Frame1.Width
        .SmallChange = 25
        .LargeChange = 100
    End With
    
    With VScroll1
        .min = 1
        .max = picMapa.Height - Frame1.Height
        .SmallChange = 25
        .LargeChange = 100
    End With

End Sub

Private Sub HScroll1_Change()

    picMapa.left = -HScroll1.Value

End Sub


Private Sub VScroll1_Change()

    picMapa.top = -VScroll1.Value

End Sub
