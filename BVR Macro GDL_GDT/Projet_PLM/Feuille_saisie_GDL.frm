VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Feuille_saisie_GDL 
   Caption         =   "Génération de gueules de loup"
   ClientHeight    =   3765
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   6010
   OleObjectBlob   =   "Feuille_saisie_GDL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Feuille_saisie_GDL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
Feuille_saisie_GDL.Hide

Dim Usel As Selection
Dim UselLB As Object
Dim InputObject(0)
Dim oStatus As String
Dim ref1, coord(2)


InputObject(0) = "Product"
Set Usel = CATIA.ActiveDocument.Selection
Set UselLB = Usel

Usel.Clear

oStatus = UselLB.SelectElement2(InputObject, "Selectionner le support", False)

If (oStatus = "Cancel") Then
Exit Sub
Else
ok_un = True
If ok_deux Then
    Bouton_ok.Enabled = True
End If
TextBox1.Text = UselLB.Item2(1).Value.PartNumber
SelRelimite = UselLB.Item2(1).Value.PartNumber

End If

Usel.Clear
Feuille_saisie_GDL.Show
End Sub

Private Sub CommandButton2_Click()
Feuille_saisie_GDL.Hide

Dim Usel As Selection
Dim UselLB As Object
Dim InputObject(0)
Dim oStatus As String
Dim ref1, coord(2)


InputObject(0) = "Product"
Set Usel = CATIA.ActiveDocument.Selection
Set UselLB = Usel


Usel.Clear

oStatus = UselLB.SelectElement2(InputObject, "Selectionner le support", False)

If (oStatus = "Cancel") Then
Exit Sub
Else

ok_deux = True
If ok_un Then
    Bouton_ok.Enabled = True
End If
TextBox2.Text = UselLB.Item2(1).Value.PartNumber
SelRelimitant = Usel.Item2(1).Value.PartNumber

End If

Usel.Clear
Feuille_saisie_GDL.Show
End Sub

Private Sub Bouton_ok_Click()
creation_GDL = "oui"
Feuille_saisie_GDL.Hide
End Sub
Private Sub Bouton_annuler_Click()
creation_GDL = "non"
fin_GDL = True
Feuille_saisie_GDL.Hide
End Sub
