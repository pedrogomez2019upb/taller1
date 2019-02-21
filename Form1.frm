VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   12765
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tercernumero 
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   3840
      Width           =   2655
   End
   Begin VB.TextBox segundonumero 
      Height          =   405
      Left            =   240
      TabIndex        =   13
      Top             =   3360
      Width           =   2655
   End
   Begin VB.TextBox primernumero 
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   2880
      Width           =   2655
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Cerrar"
      Height          =   855
      Left            =   9960
      TabIndex        =   10
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Limpiar todo"
      Height          =   975
      Left            =   9840
      TabIndex        =   9
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Calcular el mayor de tres números ( X, Y y Z)"
      Height          =   975
      Left            =   9840
      TabIndex        =   8
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Calcular cual número es mayor que otro ( X y Y )"
      Height          =   975
      Left            =   7920
      TabIndex        =   7
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Saber la división de dos números ( X/Y)"
      Height          =   975
      Left            =   7920
      TabIndex        =   6
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Saber si un numero es positivo o negativo (X)"
      Height          =   1095
      Left            =   7920
      TabIndex        =   5
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Saber cual número es par o impar (X)"
      Height          =   975
      Left            =   5880
      TabIndex        =   4
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Calcular promedio tres números (X,Y y Z)"
      Height          =   975
      Left            =   5880
      TabIndex        =   3
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox respuesta 
      Alignment       =   2  'Center
      Height          =   1575
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular promedio dos numeros ( X y Y)"
      Height          =   1095
      Left            =   5880
      TabIndex        =   0
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Por: Pedro Felipe Gómez Bonilla / 000396221"
      Height          =   495
      Left            =   8520
      TabIndex        =   17
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Valor Z"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   16
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Valor Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   14
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Valor X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Programa Multiusos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
x = primernumero.Text
y = segundonumero.Text
z = tercernumero.Text
r = (Val(x) + Val(y)) / 2
    respuesta.Text = "El promedio de los dos números ingresados es " & r
End Sub

Private Sub Command2_Click()
x = primernumero.Text
y = segundonumero.Text
z = tercernumero.Text
r = (Val(x) + Val(y) + Val(z)) / 3
    respuesta.Text = "El promedio de los tres números ingresados es " & r
End Sub

Private Sub Command3_Click()
x = primernumero.Text
y = segundonumero.Text
z = tercernumero.Text
r = (Val(x) + Val(y)) / 2
    If x Mod 2 Then
        respuesta.Text = "El número ingresado es impar"
    Else
        respuesta.Text = "El número ingresado es par"
    End If
End Sub

Private Sub Command4_Click()
x = primernumero.Text
y = segundonumero.Text
z = tercernumero.Text
    If primernumero.Text = 0 Then
        respuesta.Text = " El número ingresado es 0"
    Else
        If primernumero.Text > 0 Then
    respuesta.Text = " El número ingresado es positivo"
        Else
    respuesta.Text = "El número ingresado es negativo"
        End If
End If
End Sub

Private Sub Command5_Click()
x = primernumero.Text
y = segundonumero.Text
z = tercernumero.Text
r = (Val(x) / Val(y))
respuesta.Text = "La división de los dos números dados es de: " & r
End Sub

Private Sub Command6_Click()
x = primernumero.Text
y = segundonumero.Text
z = tercernumero.Text
If x = y Then
respuesta.Text = " Los números son iguales"
Else
If x > y Then
respuesta.Text = ("El número ") + Trim(x) + (" es mayor que ") + Trim(y)
Else
respuesta.Text = ("El número ") + Trim(y) + (" es mayor que ") + Trim(x)
End If
End If
End Sub

Private Sub Command7_Click()
x = primernumero.Text
y = segundonumero.Text
z = tercernumero.Text
If x = y = z Then
    respuesta.Text = " Los números son iguales"
Else
  If (x > y And x > z) Then
            respuesta.Text = "El numero mayor es " + Trim(x)
        Else
            If (y > x And y > z) Then
                respuesta.Text = "El numero mayor es " + Trim(y)
            Else
                respuesta.Text = "El numero mayor es " + Trim(z)
            End If
            End If
            End If
End Sub

Private Sub Command8_Click()
primernumero.Text = ""
segundonumero.Text = ""
tercernumero.Text = ""
respuesta.Text = ""
End Sub

Private Sub Command9_Click()
End

End Sub

