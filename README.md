# VBA_Creer_Image
Permet de créer une image à partir d'un graphique présent sur un onglet

##Lien vers le site
http://customoffice.github.io/VBA_Creer_Image/

## Instruction
- Soit créer un module dans votre projet vba et y copier/coller le code ci-dessous
- Soit télécharger le module (fichier *.bas) et l'inserer dans votre projet vba

##Code
```bash
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'!!!TITRE : Génère une image à partir d'un graphique                                                !!!
'!!!DATE :  17.04.15                                                                                !!!
'!!!                                                                                                !!!
'!!!DESCRIPTION : Permet de créer une image à partir d'un graphique présent sur un onglet			!!!
'!!!                                                                                                !!!
'!!!REGLES :                                                                                        !!!
'!!!- utilise le nom du graphique, le nom de l'onglet                                               !!!
'!!!- le chemin par défaut pour l'enregistrement de l'image est l'emplacement du fichier excel      !!!
'!!!- si un chemin est spécifié, par défaut il est en absolu, c'est à dire, le chemin complet, si   !!!
'!!!vous voulez utiliser le chemin en relatif, il faut forcé l'argument chemin_realtif à true       !!!
'!!!- par défaut le type d'image est png, vous pouvez spécifier jpg ou gif aussi                    !!!
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Sub creer_image(nom_graph As String, nom_feuille As String, Optional chemin As String = "", Optional chemin_relatif As Boolean = False, Optional type_image As String = "png")
    'déclaration des variables
    Dim test_creation_image As Boolean
    Dim taille As Integer
    
    'génère le chemin
    If chemin = "" Then
        chemin_graph = ThisWorkbook.Path & Application.PathSeparator & nom_graph & "." & type_image
    Else
        If chemin_relatif = False Then
            chemin_graph = chemin & Application.PathSeparator & nom_graph & "." & type_image
        Else
            chemin_graph = ThisWorkbook.Path & Application.PathSeparator & chemin & Application.PathSeparator & nom_graph & "." & type_image
        End If
    End If
    
    test_creation_image = False 'quelque fois l'enregistrement bug (pas trouvé pourquoi) cette variable permet de reboucler la création de l'image jusqu'à son bon fonctionnement
    Do Until test_creation_image = True
        Worksheets(nom_feuille).Select
        ActiveWindow.Zoom = 100 'fige le zoom car influence la taille de l'image à l'enregistrement
        Worksheets(nom_feuille).ChartObjects(nom_graph).Chart.Export _
            Filename:=chemin_graph, FilterName:=type_image
        taille = FileLen(chemin_graph)
        If taille < 100 Then
            test_creation_image = False
            Worksheets(nom_feuille).ChartObjects(nom_graph).Select
        Else
            test_creation_image = True
        End If
    Loop
End Sub
```
