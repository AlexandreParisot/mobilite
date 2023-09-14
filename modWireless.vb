Option Strict Off
Option Explicit On
Module modWireless
	
	'PREFIXE affecté au module = "WRL"
	
	'Définition des Objets nécessaires pour la communication avec le PDT
	
	Public go_AUX As New WirelessStudioOle.RFAUXPORT 'Gestion des ports COM
	Public go_BAR As New WirelessStudioOle.RFBARCODE 'Gestion du scanner
	Public go_BTN As New WirelessStudioOle.RFBUTTON 'Gestion des boutons  (Version WIN_CE exclusivement)
	Public go_ERR As New WirelessStudioOle.RFERROR 'Gestion des erreurs internes au PDT
	Public go_FIL As New WirelessStudioOle.RFFILE 'Gestion des fichiers du PDT
	Public go_IO As New WirelessStudioOle.RFIO 'Gestion de la saisie,Affichage,Impression
	Public go_LOG As New WirelessStudioOle.RFLOG 'Gestion des Logs pour le module Administrateur
	Public go_MNU As New WirelessStudioOle.RFMENU 'Gestion des menus
	Public go_TRM As New WirelessStudioOle.RFTERMINAL 'Gestion du Terminal (Config, batterie...)
	Public go_SON As New WirelessStudioOle.RFTONE 'Gestion du buzzer
	
	
	'Touches du PDT
	Public Const gCST_TOUCHE_ECHAP As Short = 27
	Public Const gCST_TOUCHE_ENTER As Short = 13
	Public Const gCST_TOUCHE_QUITTER_F3 As String = "3"
    Public Const gCST_TOUCHE_SUIVANT_F2 As String = "2"
    Public Const gCST_TOUCHE_PRECEDENT_F1 As String = "1"
    Public Const gCST_TOUCHE_MISE_EN_ATTENTE_F4 As String = "4"
    Public Const gCST_TOUCHE_PLEIN_F4 As String = "4"
    Public Const gCST_TOUCHE_PRECEDENT As String = "P"
    Public Const gCST_TOUCHE_SUIVANT As String = "S"
    Public Const gCST_TOUCHE_SAISIE_F5 As String = "5"
    Public Const gCST_TOUCHE_FIN_LIGNE As String = "6"
	Public Const gCST_TOUCHE_STOCK_ABSENT As String = "7"
	Public Const gCST_TOUCHE_STOCK_AJOUT As String = "9"
	Public Const gCST_TOUCHE_LP_SUIVANTE As String = "0"
    Public Const gCST_TOUCHE_CHANGEMENT_DE_LOT_F8 As String = "8"
    Public Const gCST_TOUCHE_IMPRIME_DERNIERE_ETIQUETTE_F2 As String = "2"

    'Fonction qui lit la configuration du terminal
    Public Function bWRL_RecupereConfiguration() As Boolean
		On Error GoTo Erreur
		bWRL_RecupereConfiguration = go_TRM.ReadTerminalInfo()
		If Len(go_TRM.TerminalID) > 0 Then
			gTab_General.sPDT = "PDT" & Mid(go_TRM.TerminalID, Len(go_TRM.TerminalID) - 2, 3)
			gTab_General.sIP_PDT = go_TRM.TerminalID
			
			WRL_InitialisationDesBoutons()
			
		End If
		Exit Function
Erreur: 
		bWRL_RecupereConfiguration = False
		LDF_LogErreurApplication(giPROC_bRecupereConfiguration)
		
	End Function
	
	'Initialisation des codes barres à utiliser sur le PDT
	'Paramètres: ( dans gTabConfiguration )
	'chaine des codes barres. On peut mettre plusieurs
	'codes barres en les séparant par le caractère "|"
	Public Function bWRL_InitCodeBarrePDT() As Boolean
		On Error GoTo Erreur
        Dim sCodeBar As String = ""
        Dim sChaine As String = ""
        Dim sChar As String = ""
        Dim n As Short
		Dim nBarCode As Short
		Dim nMin As Short
		Dim nMax As Short
		
		sCodeBar = gTab_Configuration.sCodeBar
		
		bWRL_InitCodeBarrePDT = True
		
		'Suppression des codes barres déjà configurés
		With go_BAR
			
			.ClearBarcodes()
			
			'Analyse des codes barres
			For n = 1 To Len(sCodeBar)
				
				sChar = Mid(sCodeBar, n, 1)
				
				If sChar = "|" Or n = Len(sCodeBar) Then
					If n = Len(sCodeBar) Then
						sChaine = sChaine & sChar
					End If
					If bWRL_VerifieCodeBarre(sChaine, nBarCode, nMin, nMax) Then
                        If sChaine = "CODE_128" Then
                            ' le 2ème paramètre, si on le passe à la valeur "True", permet d'activer la
                            ' prise en compte du format de CAB UCC_128 en plus de l'EAN128
                            .AddBarcode(nBarCode, True, WirelessStudioOle.RFBarcodeConstants.DECODEON, nMin, nMax)
                        Else
                            .AddBarcode(nBarCode, False, WirelessStudioOle.RFBarcodeConstants.DECODEON, nMin, nMax)
                        End If

                    Else
						MSG_AfficheErreur(giERR_CODE_BARRE, sChaine, "Invalide")
						bWRL_InitCodeBarrePDT = False
					End If
					sChaine = ""
				Else
					sChaine = sChaine & sChar
				End If
				
			Next 
			.StoreBarcode(gCST_sFICHIER_CODE_BARRE, WirelessStudioOle.RFBarcodeConstants.BCDISABLED)
		End With
		
		Exit Function
Erreur: 
		bWRL_InitCodeBarrePDT = False
		LDF_LogErreurApplication(giPROC_bInitCodeBarrePDT)
	End Function

    'Gestion des erreurs de PDT lié à la communication
    'Paramètre:
    'L'erreur retourner par le PDT via GetlastError()
    Public Sub WRL_GestionErreurPDT(ByRef vlErreur As Integer, ByRef vsSource As String)
		Dim sArg1 As String
		
		
		'Utilisé pour fermer l'application
		gbErreurCommunication = True
		
		Select Case vlErreur
			
			Case WirelessStudioOle.RFErrorConstants.WLCOMMERROR
				
			Case WirelessStudioOle.RFErrorConstants.WLFUNCTIONFAILED
				
			Case WirelessStudioOle.RFErrorConstants.WLINVALIDRETURN
				
			Case WirelessStudioOle.RFErrorConstants.WLNOTINITIALIZED
				
			Case Else
				'Ici la connexion est perdue !
				
		End Select
		sArg1 = "Connexion perdue " & vlErreur
		
		
		MSG_AfficheErreur(giERR_LIAISON_PDT, sArg1, vsSource)
		
	End Sub

    'Fonction qui vérifie le code barre passé en paramètre et retourne sa valeur
    'entière et true si le code barre est bon
    'vnMin et vnMax sont les valeurs mini et maxi de caractères admises par la norme
    Private Function bWRL_VerifieCodeBarre(ByRef vsCodeBar As String, ByRef vnBarCode As Short, ByRef vnMin As Short, ByRef vnmax As Short) As Boolean
		
		bWRL_VerifieCodeBarre = True
		
		Select Case vsCodeBar
			Case "CODE_39"
				vnBarCode = WirelessStudioOle.RFBarcodeConstants.CODE_39
				vnMin = 1
				vnmax = 32
				
			Case "UPC_A"
				vnBarCode = WirelessStudioOle.RFBarcodeConstants.UPC_A
				vnMin = 12
				vnmax = 12
				
			Case "UPC_E0"
				vnBarCode = WirelessStudioOle.RFBarcodeConstants.UPC_E0
				vnMin = 6
				vnmax = 6
				
			Case "EAN_13"
				vnBarCode = WirelessStudioOle.RFBarcodeConstants.EAN_13
				vnMin = 13
				vnmax = 13
				
			Case "EAN_8"
				vnBarCode = WirelessStudioOle.RFBarcodeConstants.EAN_8
				vnMin = 8
				vnmax = 8
				
				'Case "CODE_D25"
				'    vnBarCode = CODE_D25
				'    vnMin =
				'    vnmax =
				
			Case "CODE_I25"
				vnBarCode = WirelessStudioOle.RFBarcodeConstants.CODE_I25
				vnMin = 2
				vnmax = 14
				
			Case "CODABAR"
				vnBarCode = WirelessStudioOle.RFBarcodeConstants.CODABAR
				vnMin = 2
				vnmax = 12
				
			Case "CODE_128"
				vnBarCode = WirelessStudioOle.RFBarcodeConstants.CODE_128
				vnMin = 1
				'vnmax = 32     'V173
				vnmax = 48 'V173
				
			Case "CODE_93"
				vnBarCode = WirelessStudioOle.RFBarcodeConstants.CODE_93
				vnMin = 2
				vnmax = 14
				
				'Case "CODE_11"
				'    vnBarCode = CODE_11
				
			Case "MSI"
				vnBarCode = WirelessStudioOle.RFBarcodeConstants.MSI
				vnMin = 2
				vnmax = 8
				
			Case "UPC_E1"
				vnBarCode = WirelessStudioOle.RFBarcodeConstants.UPC_E1
				vnMin = 10
				vnmax = 10
				
				'Case "PDF_417"
				'    vnBarCode = PDF_417
				
				'Case "25_IATA"
				'    vnBarCode = D25_IATA
				
				'Case "UCC_128"
				'vnBarCode = UCC_128
				'vnMin = 1
				'vnmax = 32
				
				'Case "B_UPC"
				'    vnBarCode = B_UPC
				
			Case "TO_39"
				vnBarCode = WirelessStudioOle.RFBarcodeConstants.TO_39
				vnMin = 1
				vnmax = 1
				
				
			Case Else
				bWRL_VerifieCodeBarre = False
				
		End Select
		
	End Function
	
	'Fonction qui affiche sur la dernière ligne de l'écran le message
	'Traitement en cours..." & vsParam
	Public Sub WRL_AfficheTraitementEnCours(Optional ByRef vsParam As String = "")
		
		go_IO.RFPrint(0, go_TRM.TerminalHeight - 1, "Trt ..." & vsParam, WirelessStudioOle.RFIOConstants.WLNORMAL)
		go_IO.RFFlushoutput()
		
	End Sub
	
	Private Function WRL_InitialisationDesBoutons() As Object
		
		' 2 Columns, 1 row (top to 2 buttons in the pad ), 1 pixel border
		go_BTN.PadCreate(3, 1, 1)
		go_BTN.AddButton(0, 0, "013", "ENTRER", 0)
		go_BTN.AddButton(1, 0, "027", "CLR", 0)
		go_BTN.AddButton(2, 0, "_F3", "QUITTER", 0)
        go_BTN.PadStore(gCST_sFICHIER_BOUTONS_OK_CLR_QUIT)

        go_BTN.PadCreate(4, 2, 1)
        go_BTN.AddButton(0, 0, "013", "ENTRER", 0)
        go_BTN.AddButton(1, 0, "027", "CLR", 0)
        go_BTN.AddButton(2, 0, "_F3", "QUITTER", 0)
        go_BTN.AddButton(0, 1, "_F1", "PREC.", 0)
        go_BTN.AddButton(1, 1, "_F2", "SUIV.", 0)
        go_BTN.AddButton(2, 1, "_F5", "SAISIE", 0)
        go_BTN.AddButton(3, 1, "_F6", "FIN LIGNE", 0)
        go_BTN.PadStore(gCST_sFICHIER_BOUTONS_OK_CLR_QUIT_PRECEDENT_SUIVANT_SAISIE_FIN_LIGNE)

        go_BTN.PadCreate(4, 1, 1)
        go_BTN.AddButton(0, 0, "013", "ENTRER", 0)
        go_BTN.AddButton(1, 0, "027", "CLR", 0)
        go_BTN.AddButton(2, 0, "_F3", "QUITTER", 0)
        go_BTN.AddButton(3, 0, "_F5", "F_PAL", 0)
        go_BTN.PadStore(gCST_sFICHIER_BOUTONS_OK_CLR_QUIT_F_PAL)

        go_BTN.PadCreate(2, 1, 1)
		go_BTN.AddButton(0, 0, "013", "ENTRER", 0)
		go_BTN.AddButton(1, 0, "027", "CLR", 0)
		go_BTN.PadStore(gCST_sFICHIER_BOUTONS_OK_CLR)
		
		go_BTN.PadCreate(1, 1, 1)
		go_BTN.AddButton(0, 0, "013", "ENTRER", 0)
		go_BTN.PadStore(gCST_sFICHIER_BOUTONS_OK)

        go_BTN.PadCreate(4, 1, 1)
        go_BTN.AddButton(0, 0, "013", "ENTRER", 0)
		go_BTN.AddButton(1, 0, "027", "CLR", 0)
        go_BTN.AddButton(2, 0, "_F3", "QUITTER", 0)
        go_BTN.AddButton(3, 0, "_F5", "SAISIE", 0)
        go_BTN.PadStore(gCST_sFICHIER_BOUTONS_OK_CLR_QUIT_SAISIE)

        go_BTN.PadCreate(3, 1, 1)
		go_BTN.AddButton(0, 0, "013", "ENTRER", 0)
		go_BTN.AddButton(1, 0, "027", "CLR", 0)
		go_BTN.AddButton(2, 0, "_F3", "RETOUR", 0)
		go_BTN.PadStore(gCST_sFICHIER_BOUTONS_OK_CLR_RETOUR)
		
		go_BTN.PadCreate(3, 1, 1)
		go_BTN.AddButton(0, 0, "_F7", "ABSENT", 0)
		go_BTN.AddButton(1, 0, "_F9", "AJOUT", 0)
		go_BTN.AddButton(2, 0, "_F3", "FIN", 0)
		go_BTN.PadStore(gCST_sFICHIER_BOUTONS_FIN_ABSENT_AJOUT)
		
		go_BTN.PadCreate(2, 1, 1)
		go_BTN.AddButton(0, 0, "_F9", "AJOUT", 0)
		go_BTN.AddButton(1, 0, "_F3", "FIN", 0)
		go_BTN.PadStore(gCST_sFICHIER_BOUTONS_FIN_AJOUT)

        go_BTN.PadCreate(4, 1, 1)
        go_BTN.AddButton(0, 0, "..P", "P=PREC.", 0)
		go_BTN.AddButton(1, 0, "..S", "S=SUIV.", 0)
		go_BTN.AddButton(2, 0, "013", "ENTRER", 0)
		go_BTN.AddButton(3, 0, "027", "CLR", 0)
		go_BTN.PadStore(gCST_sFICHIER_BOUTONS_OK_CLR_PRECEDENT_SUIVANT)

        go_BTN.PadCreate(4, 1, 1)
        go_BTN.AddButton(0, 0, "_F1", "PREC.", 0)
        go_BTN.AddButton(1, 0, "_F2", "SUIV.", 0)
        go_BTN.AddButton(2, 0, "013", "ENTRER", 0)
        go_BTN.AddButton(3, 0, "_F3", "FIN", 0)
        go_BTN.PadStore(gCST_sFICHIER_BOUTONS_OK_QUIT_PRECEDENT_SUIVANT)

        go_BTN.PadCreate(3, 1, 1)
		go_BTN.AddButton(0, 0, "027", "CLR", 0)
		go_BTN.AddButton(1, 0, "_F5", "F_PAL", 0)
		go_BTN.AddButton(2, 0, "_F6", "F_LIG", 0)
		go_BTN.PadStore(gCST_sFICHIER_BOUTONS_FPAL_FLIG_CLR)
		
		go_BTN.PadCreate(2, 1, 1)
		go_BTN.AddButton(0, 0, "013", "VALIDATION", 0)
		go_BTN.AddButton(1, 0, "027", "ANNULATION", 0)
		go_BTN.PadStore(gCST_sFICHIER_BOUTONS_VALIDATION_ANNULATION)

        go_BTN.PadCreate(4, 1, 1)
        go_BTN.AddButton(0, 0, "013", "ENTRER", 0)
        go_BTN.AddButton(1, 0, "027", "CLR", 0)
        go_BTN.AddButton(2, 0, "_F2", "SUIVANT", 0)
        go_BTN.AddButton(3, 0, "_F4", "M_ATT.", 0)
        go_BTN.PadStore(gCST_sFICHIER_BOUTONS_OK_CLR_LIG_SUIVANTE_M_ATT)

        go_BTN.PadCreate(5, 1, 1)
        go_BTN.AddButton(0, 0, "013", "ENTRER", 0)
        go_BTN.AddButton(1, 0, "027", "CLR", 0)
        go_BTN.AddButton(2, 0, "_F2", "SUIVANT", 0)
        go_BTN.AddButton(3, 0, "_F4", "M_ATT.", 0)
        go_BTN.AddButton(4, 0, "_F5", "F_PAL", 0)
        go_BTN.PadStore(gCST_sFICHIER_BOUTONS_OK_CLR_LIG_SUIVANTE_M_ATT_FPAL)

        go_BTN.PadCreate(4, 1, 1)
        go_BTN.AddButton(0, 0, "027", "CLR", 0)
        go_BTN.AddButton(1, 0, "_F5", "F_PAL", 0)
        go_BTN.AddButton(2, 0, "_F6", "F_LIG", 0)
        go_BTN.AddButton(3, 0, "_F8", "CHG_LOT", 0)
        go_BTN.PadStore(gCST_sFICHIER_BOUTONS_FPAL_FLIG_CLR_CHGLOT)

        go_BTN.PadCreate(4, 1, 1)
        go_BTN.AddButton(0, 0, "013", "ENTRER", 0)
        go_BTN.AddButton(1, 0, "027", "CLR", 0)
        go_BTN.AddButton(2, 0, "_F3", "QUITTER", 0)
        go_BTN.AddButton(3, 0, "_F2", "SUIVANT", 0)
        go_BTN.PadStore(gCST_sFICHIER_BOUTONS_OK_CLR_QUIT_SUIVANT)

        go_BTN.PadCreate(4, 1, 1)
        go_BTN.AddButton(0, 0, "013", "ENTRER", 0)
        go_BTN.AddButton(1, 0, "027", "CLR", 0)
        go_BTN.AddButton(2, 0, "_F3", "QUITTER", 0)
        go_BTN.AddButton(3, 0, "_F2", "IMPR.ETIQ", 0)
        go_BTN.PadStore(gCST_sFICHIER_BOUTONS_OK_CLR_QUIT_IMPRIME_DERNIERE_ETIQUETTE)


    End Function
End Module