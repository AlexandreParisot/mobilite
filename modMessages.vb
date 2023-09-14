Option Strict Off
Option Explicit On
Module modMessages
	
	'PREFIXE affecté au module = "MSG"
	
	'Constantes d'erreurs
	Public Const giERR_OPTION_NON_PREVUE As Short = 1
	Public Const giERR_CODE_BARRE As Short = 2
	Public Const giERR_PROCEDURE_API As Short = 3
	Public Const giERR_CONFIG_INI As Short = 4
	Public Const giERR_LIAISON_PDT As Short = 5
	Public Const giERR_CONFIG_CODE_BARRE As Short = 6
    Public Const giERR_INIT_PROCEDURE_API As Short = 7
    Public Const giERR_LOGIN_USER As Short = 8
    Public Const giERR_FORMAT_NUMERIC As Short = 9

    'Droit sur menu
    Public Const giERR_INITIALISATION_LISTE_MENU As Short = 18
	Public Const giERR_DROIT_SUR_MENU As Short = 19
	Public Const giERR_INITIALISATION_MENU As Short = 20

    'Erreur sur manipulation fichier
    Public Const giERR_RECHERCHE_FICHIER As Short = 23

    'Erreur quantite à 0
    Public Const giERR_QUANTITE_NULL As Short = 24

    'ID de stock invalide
    Public Const giERR_ID_STOCK_INVALIDE As Short = 25

    'Erreur sur le Code A Barre
    Public Const giERR_CAB_INVALIDE As Short = 26
    Public Const giERR_CAB_NUM_ORDRE_INVALIDE As Short = 27

    'Erreur sur la quantité
    Public Const giERR_QTE_INVALIDE As Short = 28

    'ID de stock n'est pas au bon statut
    Public Const giERR_ID_STOCK_STATUT_INVALIDE As Short = 29

    'ID de stock non trouvé
    Public Const giERR_ID_STOCK_NON_TROUVE As Short = 30

    'Type emplacement ne correspond pas au paramétrage du GstSto.ini
    Public Const giERR_TYPE_EMPLACEMENT_INVALIDE As Short = 31

    'Code Motif inexistant
    Public Const giERR_CODE_MOTIF_INVALIDE As Short = 32

    'Erreur quantite : doit être inférieure à celle initiale
    Public Const giERR_QUANTITE_DOIT_ETRE_INFERIEURE As Short = 33

    'Le N° de LP est invalide
    Public Const giERR_LP_INVALIDE As Short = 34

    'Aucune ligne de LP à préparer
    Public Const giERR_PLUS_DE_LIGNE_DE_LP As Short = 35

    'Erreur quantite trop grande
    Public Const giERR_QUANTITE_TROP_GRANDE As Short = 36

    'Erreur Article différent de celui attendu
    Public Const giERR_ARTICLE_DIFFERENT As Short = 37

    'Erreur Article inexistante dans la LP
    Public Const giERR_ARTICLE_INEXISTANT_DANS_LP As Short = 38

    'Générique
    Public Const giERR_PREPA_SAISIE_NOMBRE_DECIMALE As Short = 98
    Public Const giERR_REPONSE_OUI_NON As Short = 99
    Public Const giERR_GENERALE As Short = 100


    'Procedure d'affichage d'une erreur sur le Terminal
    'Paramètres:
    'viErreur = N° de L'erreur
    Public Sub MSG_AfficheErreur(ByRef viErreur As Short, Optional ByRef vsArg1 As String = "", Optional ByRef vsArg2 As String = "")
        Dim sLigne As String = ""
        Dim sLigne2 As String = ""
        Dim sLigne3 As String = ""
        Dim sLigne4 As String = ""
        Dim sLigne5 As String = ""
        Dim sLigne6 As String = ""
        Dim nLEncours As Short
        Dim n As Short

        For n = 0 To (UBound(gTab_Messages) - 1)
            gTab_Messages(n) = ""
        Next

        Select Case viErreur
            Case giERR_OPTION_NON_PREVUE
                sLigne = "Option non prévue"

            Case giERR_CODE_BARRE
                sLigne = "Erreur Code Barre"

            Case giERR_INIT_PROCEDURE_API
                sLigne = "Erreur API:"

            Case giERR_PROCEDURE_API
                sLigne = "Message"

            Case giERR_CONFIG_INI
                sLigne = "Erreur .Ini"

            Case giERR_LOGIN_USER
                sLigne = "Identification"
                vsArg1 = "Incorrecte."

            Case giERR_FORMAT_NUMERIC
                sLigne = "Valeur numérique "
                vsArg1 = "Incorrecte."

            Case giERR_LIAISON_PDT
                sLigne = "Erreur sock PDT:"

            Case giERR_CONFIG_CODE_BARRE
                sLigne = "Erreur cfg CBAR"

            Case giERR_DROIT_SUR_MENU
                sLigne = "Erreur init. menu"

            Case giERR_GENERALE
                sLigne = "Erreur..."

            Case giERR_RECHERCHE_FICHIER
                sLigne = "Erreur Fichier"

            Case giERR_REPONSE_OUI_NON
                sLigne = "Repondre par Oui ou Non."

            Case giERR_PREPA_SAISIE_NOMBRE_DECIMALE
                sLigne = "Nombre de decimales incorrect"

            Case giERR_QUANTITE_NULL
                sLigne = "Erreur: Quantite = 0"

            Case giERR_ID_STOCK_INVALIDE
                sLigne = "ID de stock inexistant"

            Case giERR_ID_STOCK_NON_TROUVE
                sLigne = "ID de stock non trouvé"

            Case giERR_CAB_INVALIDE
                sLigne = "Code à barre invalide"

            Case giERR_CAB_NUM_ORDRE_INVALIDE
                sLigne = "Code à barre"
                vsArg1 = "Numero ordre invalide"
                vsArg2 = "pour ce lot"

            Case giERR_QTE_INVALIDE
                sLigne = "Quantité invalide"
                vsArg1 = "Pas assez en stock"

            Case giERR_ID_STOCK_STATUT_INVALIDE
                sLigne = "ID de stock"
                vsArg1 = "Statut " & vsArg1 & " invalide"
                vsArg2 = "Doit être au statut " & vsArg2

            Case giERR_TYPE_EMPLACEMENT_INVALIDE
                sLigne = "Type emplacement invalide"

            Case giERR_CODE_MOTIF_INVALIDE
                sLigne = "Code motif invalide"

            Case giERR_QUANTITE_DOIT_ETRE_INFERIEURE
                sLigne = "Quantité invalide"
                vsArg1 = "Doit être inférieure"
                vsArg2 = "à celle initiale"

            Case giERR_LP_INVALIDE
                sLigne = "LP invalide"

            Case giERR_PLUS_DE_LIGNE_DE_LP
                sLigne = "Plus de ligne de LP"
                vsArg1 = "à préparer"

            Case giERR_QUANTITE_TROP_GRANDE
                sLigne = "Quantité trop grande"
                vsArg1 = "par rapport au restant"
                vsArg2 = "à préparer"

            Case giERR_ARTICLE_DIFFERENT
                sLigne = "Article différent"
                vsArg1 = "de celui attendu"

            Case giERR_ARTICLE_INEXISTANT_DANS_LP
                sLigne = "Article inexistant"
                vsArg1 = "dans la LP"

            Case Else
                sLigne = "Erreur inconnue"
				
		End Select
		
		LDF_AfficheErreurDansLog(sLigne, vsArg1, vsArg2)
		
		MSG_VerifieTailleArguments(vsArg1, sLigne2, sLigne3, sLigne4, sLigne5, sLigne6, nLEncours)
		
		MSG_VerifieTailleArguments(vsArg2, sLigne2, sLigne3, sLigne4, sLigne5, sLigne6, nLEncours)
		
		
		With go_ERR
			.ClearError()
			
			.SetErrorLine(sLigne, 0)
			For n = 1 To (UBound(gTab_Messages) - 1)
				If gTab_Messages(n) <> "" Then
					.SetErrorLine(gTab_Messages(n), n)
				End If
			Next 
			.Display(gTab_Configuration.iDelaiMessage)
			.ClearError()
		End With
		
	End Sub
	
	'Verifie la taille des arguments si le premier dépasse la largeur de l'écran
	'la suite est basculée dans le deuxième
	Private Sub MSG_VerifieTailleArguments(ByRef vsLigne1 As String, ByRef vsLigne2 As String, ByRef vsLigne3 As String, ByRef vsLigne4 As String, ByRef vsLigne5 As String, ByRef vsLigne6 As String, ByRef vnLEncOurs As Short)
		On Error GoTo Erreur
		
		Dim nTailleEcran As Short
		Dim n As Short
		Dim nDiv As Short
		Dim bAtt As Boolean
		
		
		If Trim(Mid(vsLigne1, 1, 3)) = "OK" Or Trim(Mid(vsLigne1, 1, 3)) = "NOK" Then
			vsLigne1 = Trim(Mid(vsLigne1, 4, Len(vsLigne1)))
		End If
		
		nTailleEcran = go_TRM.TerminalWidth - 3
		If (1 + vnLEncOurs) <= 12 Then
			gTab_Messages(1 + vnLEncOurs) = vsLigne1
			If nTailleEcran > 0 Then
				nDiv = (Len(vsLigne1) \ nTailleEcran) + 1
				If nDiv >= 2 Then
					If nDiv + vnLEncOurs >= 6 Then
						gTab_Messages(6 + vnLEncOurs) = Trim(Mid(vsLigne1, 5 * nTailleEcran, nTailleEcran))
					End If
					If nDiv + vnLEncOurs >= 5 Then
						gTab_Messages(5 + vnLEncOurs) = Trim(Mid(vsLigne1, 4 * nTailleEcran, nTailleEcran))
					End If
					If nDiv + vnLEncOurs >= 4 Then
						gTab_Messages(4 + vnLEncOurs) = Trim(Mid(vsLigne1, 3 * nTailleEcran, nTailleEcran))
					End If
					If nDiv + vnLEncOurs >= 3 Then
						gTab_Messages(3 + vnLEncOurs) = Trim(Mid(vsLigne1, 2 * nTailleEcran, nTailleEcran))
					End If
					If nDiv + vnLEncOurs >= 2 Then
						gTab_Messages(2 + vnLEncOurs) = Trim(Mid(vsLigne1, nTailleEcran, nTailleEcran))
					End If
					gTab_Messages(1 + vnLEncOurs) = Trim(Mid(vsLigne1, 1, nTailleEcran))
				End If
			End If
		End If
		vnLEncOurs = vnLEncOurs + nDiv
		Exit Sub
Erreur: 
		MSG_AfficheErreur(giERR_OPTION_NON_PREVUE, "Erreur MSG PDT")
		
	End Sub
End Module