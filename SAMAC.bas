Attribute VB_Name = "Module1"
'Nom : SAMAC, ou SAP Automatic Material Creation Macro
'Auteur : Sami Nouidri (SNO)
'Date : 09.07.2021


'Importation de la fonction Sleep() de windows, pr�sente dans la librairie kernel32
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



'Importation des fonctions de gestions de processus de windows, pr�sente dans la librairie user32
'Il est obligatoire d'inclure le call "PtrSafe" pour les syst�mes 64-Bit x86
Private Declare PtrSafe Function GetDesktopWindow _
  Lib "USER32.dll" () As Long
  
Private Declare PtrSafe Function GetWindow _
  Lib "USER32.dll" _
    (ByVal hWnd As Long, _
     ByVal wCmd As Long) As Long
      
Private Declare PtrSafe Function GetWindowText _
  Lib "USER32.dll" _
    Alias "GetWindowTextA" _
      (ByVal hWnd As Long, _
       ByVal lpSting As String, _
       ByVal nMaxCount As Long) As Long

'Main
Sub SAP()

'Raccourci de commande SHELL
Set x = CreateObject("wscript.shell")



Dim lRow As Long
Dim lRow_UNIV As Long
Dim Designation As String
Dim DesignationEN As String
Dim TypeArt As String
Dim GrpMarch As String
Dim AncienNum As String
Dim Labo As String
Dim CatPasan As String
Dim Fournisseur As String
Dim NumFournisseur As String
Dim NumFabricant As String
Dim TexteCommandeAchat As String
Dim DataError As Boolean
Dim Dash As Integer
Dim FirstPass As Boolean
Dim TexteCommandeAchatFlag As Boolean

Dim TimeToCopy As Integer
Dim SAPcheck As Integer
Dim ArticleCheck As Integer
Dim EndCheck As Integer
Dim MatriculeCheck As Integer
Dim Time As Long
Dim StartDelay As Long


Dim lRow_RA_CL As Long
Dim lRow_NA_CL As Long
Dim NB_articles As Long

Dim i As Integer

'objet utilis� pour le presse papier windows
Dim objData As New MSForms.DataObject
Dim NumArticleCR As String

Dim ArticlesCR As String

If ActiveWorkbook.ReadOnly Then
 MsgBox "Le classeur excel est en lecture seule, le comportement du script n'est pas garantie sous ces conditions. L'injection, et la sauveguarde du num�ro d'article cr�e est impossible.", vbExclamation
 End
End If



'Derniere ligne excel ou se trouve un num�ro d'article + 1 = la ligne ou se trouve le prochain article que l'on veut cr�e
lRow_NA_CL = Worksheets("RES_NUM_SAP").Cells(Worksheets("RES_NUM_SAP").Rows.Count, 2).End(xlUp).Row
'Derniere ligne ou se trouve une matricule : definit le nombre d'articles � creer
lRow_RA_CL = Worksheets("RES_NUM_SAP").Cells(Worksheets("RES_NUM_SAP").Rows.Count, 3).End(xlUp).Row

NB_articles = lRow_RA_CL - lRow_NA_CL

If NB_articles = 0 Then
    NB_articles = 1
ElseIf NB_articles < 0 Then
    MsgBox "Erreur lors du calcul du nombre de num�ros d'articles, veuillez verifier que l'avant derni�re ligne contient bien une matricule et un num�ro d'article.", vbCritical
    End
End If

'On part du pricipe qu'il n'y a pas de texte commande achat, a moins qu'il soit detect� lors de la v�rification
TexteCommandeAchatFlag = False


'Premier passage pour le MsgBox d'instance SAP
FirstPass = True

i = 0

'Boucle continue pour cr�er tout les articles
For i = 1 To NB_articles

'N�cessaire pour effacer le presse papier avant l'execution
Application.CutCopyMode = False

lRow_UNIV = (lRow_NA_CL + i)
DataError = False

'R�cuperation des donn�es du fichier RES_NUM
Designation = Cells(lRow_UNIV, 4).Value

If Right$(Designation, 1) = "%" Then
    Designation = Left$(Designation, Len(Designation) - 1) & " PERCENT"
End If

DesignationEN = Cells(lRow_UNIV, 5).Value

If Right$(DesignationEN, 1) = "%" Then
    DesignationEN = Left$(DesignationEN, Len(DesignationEN) - 1) & " PERCENT"
End If


TypeArt = Cells(lRow_UNIV, 6).Value
GrpMarch = Cells(lRow_UNIV, 14).Value
AncienNum = Cells(lRow_UNIV, 8).Value
Labo = Cells(lRow_UNIV, 7).Value
CatPasan = Cells(lRow_UNIV, 9).Value
Fournisseur = Cells(lRow_UNIV, 10).Value
NumFournisseur = Cells(lRow_UNIV, 11).Value
NumFabricant = Cells(lRow_UNIV, 12).Value
TexteCommandeAchat = Cells(lRow_UNIV, 13).Value


'Verification des informations initiales (voir page confluence pour plus de d�tails)
If Designation = "-" Or Len(Designation) <= 0 Or Designation = Null Then
    MsgBox "Erreur lors de la r�cuperation de donn�e : Manque une d�signation � la ligne " & CStr(lRow_UNIV), vbCritical
    DataError = True
End If
If DesignationEN = "-" Or Len(DesignationEN) <= 0 Or DesignationEN = Null Then
    MsgBox "Erreur lors de la r�cuperation de donn�e : Manque une d�signation anglaise � la ligne " & CStr(lRow_UNIV), vbCritical
    DataError = True
End If
If TypeArt = "-" Or Len(TypeArt) <= 0 Or TypeArt = Null Then
    MsgBox "Erreur lors de la r�cuperation de donn�e : Manque une unit� de base � la ligne " & CStr(lRow_UNIV), vbCritical
    DataError = True
End If
If Labo = "-" Or Len(Labo) <= 0 Or Labo = Null Then
    MsgBox "Erreur lors de la r�cuperation de donn�e : Manque un Labo/Bur. d'�tudes � la ligne " & CStr(lRow_UNIV), vbCritical
    DataError = True
End If
If AncienNum = "-" Or Len(AncienNum) <= 0 Or AncienNum = Null Then
    MsgBox "Erreur lors de la r�cuperation de donn�e : Manque un ancien num�ro d'article � la ligne " & CStr(lRow_UNIV), vbCritical
    DataError = True
End If
If GrpMarch = "-" Or Len(GrpMarch) <= 0 Or GrpMarch = Null Then
    MsgBox "Erreur lors de la r�cuperation de donn�e : Manque un groupe marchand � la ligne " & CStr(lRow_UNIV), vbCritical
    DataError = True
End If
If CatPasan = "-" Or Len(CatPasan) <= 0 Or CatPasan = Null Then
    MsgBox "Erreur lors de la r�cuperation de donn�e : Manque une cat�gorie de classification � la ligne " & CStr(lRow_UNIV), vbCritical
    DataError = True
End If
If Fournisseur = "-" Or Len(Fournisseur) <= 0 Or Fournisseur = Null Then
    MsgBox "Erreur lors de la r�cuperation de donn�e : Manque un fournisseur � la ligne " & CStr(lRow_UNIV) & ", Veuillez renseigner le champ avec la valeur 'Pasan' au minimum.", vbCritical
    DataError = True
End If
If Len(NumFournisseur) <= 0 Or NumFournisseur = Null Then
    MsgBox "Erreur lors de la r�cuperation de donn�e : Manque un num�ro de fournisseur � la ligne " & CStr(lRow_UNIV) & ", Veuillez renseigner le champ par un '-' au minimum.", vbCritical
    DataError = True
End If
If Len(NumFabricant) <= 0 Or NumFabricant = Null Then
    MsgBox "Erreur lors de la r�cuperation de donn�e : Manque un fabricant � la ligne " & CStr(lRow_UNIV) & ", Veuillez renseigner le champ par un '-' au minimum.", vbCritical
    DataError = True
End If
If Len(TexteCommandeAchat) <= 0 Or TexteCommandeAchat = Null Then
    MsgBox "Erreur lors de la r�cuperation de donn�e : Manque un texte commande achat � la ligne " & CStr(lRow_UNIV) & ", Veuillez renseigner le champ par un '-' au minimum.", vbCritical
    DataError = True
End If


If DataError = True Then
    End
    'erreur de r�cuperation, le programe se ferme
End If

'Verifier si il y a un texte commande achat, si oui, avertir � la fin qu'il n'est pas pris en compte
Dash = StrComp(TexteCommandeAchat, "-")
If Len(TexteCommandeAchat) > 0 And Dash <> 0 Then
    TexteCommandeAchatFlag = True
End If


'Avertissement avant de fermer SAP
If FirstPass = True Then

    SAPcheck = MsgBox("Si une instance de SAP est ouverte, ce script va la fermer, sans sauvegarder. Voulez-vous continuer?", vbCritical + vbYesNo)

    If SAPcheck = vbNo Then
        End
    End If

End If

EXEC:
'fermeture de toute instance de SAP
Shell ("taskkill.exe /f /t /im saplgpad.exe")
Shell ("taskkill.exe /f /t /im SAPgui.exe")
StartDelay = 0
Do
    Debug.Print ("Closing SAP...")
    Debug.Print (StartDelay)
    Sleep (1)
    StartDelay = StartDelay + 1
Loop Until IsProcessRunning("saplgpad.exe") = False And IsProcessRunning("SAPgui.exe") = False Or StartDelay > 500
If StartDelay > 500 Then
    MsgBox "Erreur lors de la fermeture de SAP.", vbCritical
    End
Else
    Debug.Print ("Exec")
End If

x.Run "MM01.SAP"
StartDelay = 0
Do
    Debug.Print ("Opening SAP...")
    Debug.Print (StartDelay)
    Sleep (1)
    StartDelay = StartDelay + 1
Loop Until IsWindowOpen("Cr�er Article (Ecran initial)") = True Or StartDelay > 500
If StartDelay > 500 Then
    MsgBox "Aucune Instance de SAP detect�e", vbCritical
    End
Else
    Debug.Print ("Exec")
End If
Debug.Print ActiveWindow.Caption
Sleep 500
'Ecran Initial
x.SendKeys "{TAB}"
Sleep 250
x.SendKeys "Mec"
Sleep 500
x.SendKeys "{TAB}"
Sleep 500
x.SendKeys "M+"
Sleep 1000
x.SendKeys "{enter}"
StartDelay = 0
Do
    Debug.Print ("Opening Material creation...")
    Debug.Print (StartDelay)
    Sleep (1)
    StartDelay = StartDelay + 1
    If StartDelay = 50 Then
        x.SendKeys "{enter}"
        Sleep 500
        Debug.Print ("Failed last attempt")
    End If
Loop Until IsWindowOpen("Cr�er Article (Ecran initial)") = False Or StartDelay > 500
If StartDelay > 500 Then
    MsgBox "Erreur lors l'acc�s � la cr�ation d'articles initiales, SAP timeout.", vbCritical
    End
Else
    Debug.Print ("Exec")
End If

'Cases obligatoires
x.SendKeys (Designation)
Sleep 1000
x.SendKeys "{TAB}"
Sleep 250
x.SendKeys "{TAB}"
Sleep 250
x.SendKeys "{TAB}"
Sleep 250
x.SendKeys (TypeArt)
Sleep 500
x.SendKeys "{TAB}"
Sleep 250
x.SendKeys (GrpMarch)
Sleep 1000
x.SendKeys "{TAB}"
Sleep 250
x.SendKeys (AncienNum)
Sleep 500
x.SendKeys "{TAB}"
Sleep 250
x.SendKeys "{TAB}"
Sleep 250
x.SendKeys "{TAB}"
Sleep 250
x.SendKeys (Labo)
Sleep 500
x.SendKeys "{enter}"
Sleep 500
x.SendKeys "{enter}"
Sleep 500
x.SendKeys "{enter}"
Sleep 500
'Classification


StartDelay = 0
Do
    Debug.Print ("Opening Second Classification...")
    Debug.Print (StartDelay)
    Sleep (1)
    StartDelay = StartDelay + 1
    If StartDelay = 50 Then
        x.SendKeys "{enter}"
        Sleep 500
        Debug.Print ("Failed last attempt")
    End If
Loop Until IsWindowOpen("classification") = True Or StartDelay > 500

If StartDelay > 500 Then
    MsgBox "Erreur lors l'acc�s � la classification, SAP timeout.", vbCritical
    End
Else
    Debug.Print ("Exec")
End If
x.SendKeys "mbmaterial"
Sleep 500
x.SendKeys "{enter}"
Sleep 500
x.SendKeys (CatPasan)
Sleep 500
x.SendKeys "{TAB}"
Sleep 250
x.SendKeys "{TAB}"
Sleep 250
x.SendKeys "{TAB}"
Sleep 250
x.SendKeys "{TAB}"
Sleep 250
x.SendKeys (Fournisseur)
Sleep 500
x.SendKeys "{TAB}"
Sleep 250
x.SendKeys (NumFournisseur)
Sleep 500
x.SendKeys "{TAB}"
Sleep 250
x.SendKeys (NumFabricant)
Sleep 500
x.SendKeys "{enter}"
Sleep 500
x.SendKeys "{enter}"
Sleep 500
x.SendKeys "^{F8}"

'Designation EN
Sleep 1000
x.SendKeys "^{F6}"
Sleep 1000
x.SendKeys "{TAB}"
Sleep 250
x.SendKeys "{TAB}"
Sleep 250
x.SendKeys "EN"
Sleep 500
x.SendKeys "{TAB}"
Sleep 250
x.SendKeys (DesignationEN)
Sleep 750
x.SendKeys "^{F3}"
Sleep 500

    

'Sauveguarde du numero cr�e
x.SendKeys "{F3}"
Sleep 250
x.SendKeys "{F3}"
Sleep 250
x.SendKeys "{enter}"
StartDelay = 0
Do
    Debug.Print ("Going back to initial screen...")
    Debug.Print (StartDelay)
    Sleep (1)
    StartDelay = StartDelay + 1
    If StartDelay = 50 Then
        x.SendKeys "{enter}"
        Sleep 500
        Debug.Print ("Failed last attempt")
    End If
Loop Until IsWindowOpen("Cr�er Article (Ecran initial)") = True Or StartDelay > 500
If StartDelay > 500 Then
    MsgBox "Erreur lors l'acc�s � l'�cran initiale, SAP timeout.", vbCritical
    End
Else
    Debug.Print ("Exec")
End If

'Reprise du num�ro
x.SendKeys "{TAB}"
Sleep 250
x.SendKeys "{TAB}"
Sleep 250
x.SendKeys "{TAB}"
Sleep 250
x.SendKeys "{TAB}"
Sleep 250
x.SendKeys "{TAB}"
Sleep 250
x.SendKeys "{TAB}"
Sleep 250
x.SendKeys "/nmm03"
Sleep 250
x.SendKeys "{enter}"
StartDelay = 0
Do
    Debug.Print ("Opening MM03...")
    Debug.Print (StartDelay)
    Sleep (1)
    StartDelay = StartDelay + 1
    If StartDelay = 50 Then
        x.SendKeys "{enter}"
        Sleep 500
        Debug.Print ("Failed last attempt")
    End If
Loop Until IsWindowOpen("Afficher Article (Ecran initial)") = True Or StartDelay > 500
If StartDelay > 500 Then
    MsgBox "Erreur lors l'acc�s � la MM03, SAP timeout.", vbCritical
    End
Else
    Debug.Print ("Exec")
End If
Debug.Print ("Entering ctrl loop")



TimeToCopy = 0
Do
x.SendKeys "^a"
Sleep 1000
x.SendKeys "^c"
Sleep 1000
Debug.Print ("CTRL Loop")
'envoi vers excel

objData.GetFromClipboard
DoEvents
NumArticleCR = objData.GetText()
DoEvents
TimeToCopy = TimeToCopy + 1

Debug.Print (NumArticleCR)
Debug.Print (TimeToCopy)

Loop Until Len(NumArticleCR) > 0

Debug.Print ("Exiting ctrl loop ")

'ecriture dans la bonne celllule
Worksheets("RES_NUM_SAP").Range("B" & lRow_UNIV).Value = NumArticleCR
Sleep 1000

'Message de fin pour le mode manuel
If Worksheets("RES_NUM_SAP").CheckBox1.Value = False Then
'si il y avait un texte commande chat, avertir qu'il n'as pas �t� pris en compte
    If TexteCommandeAchatFlag = True Then
        MsgBox "Un ou Plusieurs articles cr�e(s) poss�dent un texte commande achat. Veuillez notez que le texte commande achat n'est pas pris en compte lors de la cr�ation automatique de l'article, il doit etre ajout� manuellement par la suite.", vbInformation
    End If
    MsgBox "Article " & NumArticleCR & " cr�e", vbInformation
    End
End If

'On passe a la suite
FirstPass = False
Debug.Print (i)
ArticlesCR = ArticlesCR + NumArticleCR + ", "
Next i

'si il y avait un texte commande chat, avertir qu'il n'as pas �t� pris en compte
If TexteCommandeAchatFlag = True Then
        MsgBox "Un ou Plusieurs articles cr�e(s) poss�dent un texte commande achat. Veuillez notez que le texte commande achat n'est pas pris en compte lors de la cr�ation automatique de l'article, il doit etre ajout� manuellement par la suite.", vbInformation
End If
'Message de fin pour le mode automatique
MsgBox "Les articles " & ArticlesCR & " ont �t� cr�es.", vbInformation


End Sub

Function IsProcessRunning(process As String)
'written : July 21, 2017
'author : enderland
'summary : checks if a process or executable is running, returns true if it does
'source : https://stackoverflow.com/questions/29807691/determine-if-application-is-running-with-excel
    Dim objList As Object

    Set objList = GetObject("winmgmts:") _
        .ExecQuery("select * from win32_process where name='" & process & "'")

    If objList.Count > 0 Then
        IsProcessRunning = True
    Else
        IsProcessRunning = False
    End If

End Function

Public Function IsWindowOpen(ByVal Window_Caption As String) As Boolean

'Written: October 12, 2011
'Author:  Leith Ross
'Summary: Compares the supplied window caption against all top level windows currently open.
'         The window caption will match if complete or partial. Case is ignored.
'Source : https://www.mrexcel.com/board/threads/check-if-application-open-before-activating.585274/

    Dim Caption As String
    Dim CurrWnd As Long
    Dim L As Long
    
    Const GW_CHILD As Long = 5
    Const GW_HWNDNEXT As Long = 2
  
  
       ' Start with the Top most window that has the focus
         CurrWnd = GetWindow(GetDesktopWindow, GW_CHILD)
      
          ' Loop while the hWnd returned by GetWindow is valid.
            While CurrWnd <> 0
         
              ' Get Window caption
                Caption = String(64, Chr$(0))
                L = GetWindowText(CurrWnd, Caption, 64)
                Caption = IIf(L > 0, Left(Caption, L), "")
         
              ' Test if the caption matches the Window requested
                If LCase(Caption) Like "*" & LCase(Window_Caption) & "*" Then
                   IsWindowOpen = True
                   Exit Function
                End If
         
              ' Get the next Window
                CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
         
              ' Process Windows events.
                DoEvents
            Wend
    
End Function










