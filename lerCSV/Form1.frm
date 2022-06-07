VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4740
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command 
      Caption         =   "Ler Arquivo"
      Height          =   615
      Left            =   5400
      TabIndex        =   0
      Top             =   4080
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command_Click()

F = FreeFile
Open "ncm.csv" For Input As F
Line Input #F, linha
arq = Split(linha, vbLf)
Close #F


For a = 0 To UBound(arq)
   
   Print a & "/" & UBound(arq)
   linha = arq(a)
   x = ""
   
   For i = 1 To Len(linha)
       If Mid(linha, i, 1) <> """" Then
          x = x & Mid(linha, i, 1)
       End If
   Next
   
   Debug.Print x & " ."
Next

Close #F
End Sub
