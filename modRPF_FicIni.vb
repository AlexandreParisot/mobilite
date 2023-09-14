Option Strict Off
Option Explicit On
Module modRPF_FicIni
	
	'PREFIXE affecté au module = "INI"
	Public Function bINI_InitParametrage() As Boolean
		
		If bINI_Init_GstSto() Then
            If bINI_Init_Menu() Then
                bINI_InitParametrage = True
            End If
        End If
		
	End Function

    ' Retourne la chaine contenu dans un fichier INI
    ' Argument:
    '   - nom de section
    '   - nom de clef
    ' Retourne "" si chaine vide ou resultat non trouvé
    Public Function sINI_GetChaineFichierIni(ByRef vsSection As String, ByRef vsClef As String, ByRef vsFicIni As String) As String

        Dim sResul As String
        Dim longueur As Integer
        sINI_GetChaineFichierIni = ""


        Const LG_MAX_CHAINE As Integer = 600

        longueur = LG_MAX_CHAINE
        sResul = New String(Chr(0), longueur)

        longueur = GetPrivateProfileString(vsSection, vsClef, "", sResul, longueur, vsFicIni)
        If longueur > 0 Then
            sINI_GetChaineFichierIni = Left(sResul, longueur)
        End If
    End Function


    'Fonction qui renseigne la structure gTab_Configuration à partir du fichier Ini de paramètrage
    'Renvoie True si les champs obligatoires ont été renseignés
    Public Function bINI_Init_GstSto() As Boolean
		On Error GoTo Erreur
        Dim sFicIni As String = ""
        Dim sBool As String = ""
        Dim bFin As Boolean
		Dim nUser As Short
        Dim sUSER As String = ""
        Dim bErrUser As Boolean
        Dim sErreur As String = ""

        sFicIni = My.Application.Info.DirectoryPath & "\" & gCST_sFICHIER_INI
		
		
		With gTab_Configuration

            'Section API
            .sIP = sINI_GetChaineFichierIni(gCST_INI_SEC_API, gCST_INI_IP, sFicIni)
            .sPort = sINI_GetChaineFichierIni(gCST_INI_SEC_API, gCST_INI_PORT, sFicIni)
            .sDomaine = sINI_GetChaineFichierIni(gCST_INI_SEC_API, gCST_INI_DOMAINE, sFicIni)

            'Section PDT
            .iDelaiMessage = Val(sINI_GetChaineFichierIni(gCST_INI_SEC_PDT, gCST_INI_DLA, sFicIni))
			.sCodeBar = sINI_GetChaineFichierIni(gCST_INI_SEC_PDT, gCST_INI_CODBAR, sFicIni)

            'Section M3
            .sSociete = sINI_GetChaineFichierIni(gCST_INI_SEC_M3, gCST_INI_CONO, sFicIni)
            .sDivision = sINI_GetChaineFichierIni(gCST_INI_SEC_M3, gCST_INI_DIVI, sFicIni)
            .sDepot = sINI_GetChaineFichierIni(gCST_INI_SEC_M3, gCST_INI_WHLO, sFicIni)
            .sSLPT_Bord_De_Chaine_Normal = sINI_GetChaineFichierIni(gCST_INI_SEC_M3, gCST_INI_TYPE_EMPLACEMENT_BORD_DE_CHAINE_NORMAL, sFicIni)
            .sSLPT_Bord_De_Chaine_Bib = sINI_GetChaineFichierIni(gCST_INI_SEC_M3, gCST_INI_TYPE_EMPLACEMENT_BORD_DE_CHAINE_BIB, sFicIni)

            'Section APP
            sBool = sINI_GetChaineFichierIni(gCST_INI_SEC_APP, gCST_INI_LOG, sFicIni)
			If UCase(sBool) = "1" Or UCase(sBool) = "0" Then
				.bLog = CBool(sBool)
			End If
            .lTimeWait = Val(sINI_GetChaineFichierIni(gCST_INI_SEC_APP, gCST_INI_TIMW, sFicIni))

            'Section EEP

            'Section ABC

            'Section RES
            .sRES_WHSL_TB_Final = sINI_GetChaineFichierIni(gCST_INI_SEC_RES, gCST_INI_RES_EMPLACEMENT_FINAL_POUR_TB, sFicIni)

            'Section TDS

            'Section EMS

            'Section TID

            'Section USER
            'Memorisation des utilisateurs paramétrés
            ReDim Preserve .sProfil(0)
			While Not bFin
				nUser = nUser + 1
				sUSER = sINI_GetChaineFichierIni(gCST_INI_SEC_USER, CStr(nUser), sFicIni)
				If sUSER <> "" Then
                    ReDim Preserve .sProfil(nUser)
                    .sProfil(nUser).sUtilisateur = sUSER
                    If .sProfil(nUser).sUtilisateur = "" Then
                        bErrUser = True
                    End If
                Else
					bFin = True
				End If
			End While



            'Vérification des paramètres obligatoires
            'L'utilisateur et le mot de passe de la section "API" ne sont plus obligatoires

            If .sIP = "" Then
                sErreur = "IP du serveur M3BE n'est pas renseigne."
            End If
            If .sPort = "" Then
                sErreur = "Le Port d'acces du serveur M3BE n'est pas renseigne pour ce PDT."
            End If
            If .sDomaine = "" Then
                sErreur = "Le Domaine réseau n'est pas renseigne."
            End If
            If .sSociete = "" Then
				sErreur = "La Societe n'est pas renseignee."
			End If
            If .sDivision = "" Then
                sErreur = "La Division n'est pas renseignee."
            End If
            If .sDepot = "" Then
                sErreur = "Le depot n'est pas renseignee."
            End If
            If .sCodeBar = "" Then
				sErreur = "Aucun code a barre n'est renseigne."
			End If
            If .lTimeWait <= 0 Then
                sErreur = "La valeur du délai de connexion ne peut-etre negative."
            End If
            If .sSLPT_Bord_De_Chaine_Normal = "" Then
                sErreur = "Type emplacement bord de chaine normal n'est pas renseigne."
            End If
            If .sSLPT_Bord_De_Chaine_Bib = "" Then
                sErreur = "Type emplacement bord de chaine bib n'est pas renseigne."
            End If
            If .sRES_WHSL_TB_Final = "" Then
                sErreur = "RES - Emplacement de retour en stock pour les TB n'est pas renseigne."
            End If
            If bErrUser Then
                sErreur = "Une erreur s'est produite dans la configuration utilisateur."
            End If

            If sErreur <> "" Then
				MSG_AfficheErreur(giERR_CONFIG_INI, "Parametrage .ini incomplet", sErreur)
			Else
				bINI_Init_GstSto = True
			End If
			
		End With
		Exit Function
Erreur: 
		MSG_AfficheErreur(giERR_CONFIG_INI, "Erreur pendant l' initialisation des parametrages : " & Err.Number & "=>", Err.Description)
	End Function

    'Fonction de lecture du fichier Menu.ini
    'Lecture des Menus Paramètrés
    'Les valeurs sont stockées dans gTab_Menu
    Public Function bINI_Init_Menu() As Boolean
		On Error GoTo Erreur
		
		Dim nNumMenu As Short
		Dim sMenu As String
		Dim sFicMenuIni As String
		
		sFicMenuIni = My.Application.Info.DirectoryPath & "\" & gCST_sFICHIER_MENU_INI
		
		With gTab_Menu
			
			ReDim Preserve .Tab_Menu(0)
			
			'Initialisation de la liste des menus possibles
			For nNumMenu = 1 To 100
				sMenu = sINI_GetChaineFichierIni(gCST_INI_SEC_MENU, "OPT_" & nNumMenu, sFicMenuIni)
				If sMenu <> "" Then
					.nNombreMenu = .nNombreMenu + 1
					ReDim Preserve .Tab_Menu(.nNombreMenu)
                    .Tab_Menu(.nNombreMenu) = sMenu
                End If
			Next 
		End With
		
		bINI_Init_Menu = True
		Exit Function
Erreur: 
		MSG_AfficheErreur(giERR_INITIALISATION_LISTE_MENU, "Erreur pendant initialisation de la liste des menus : " & Err.Number & "=>", Err.Description)
	End Function

    'Fonction de lecture du fichier Menu.ini contenant les droits utilisateurs
    'Les valeurs sont stockées dans vobjTabMenu
    Public Function bINI_Init_DroitSurMenu(ByRef vsUtilisateur As String) As Boolean
		On Error GoTo Erreur
		
		Dim nNumMenu As Short
		Dim sBool As String
		Dim sFicMenuIni As String
		
		sFicMenuIni = My.Application.Info.DirectoryPath & "\" & gCST_sFICHIER_MENU_INI
		
		With gTab_Menu
			
			ReDim .Tab_DroitMenu(.nNombreMenu)
			'Initialisation de la liste des menus possibles
			For nNumMenu = 1 To .nNombreMenu
				'Gestion de la securite
				If bSEC_AnalyseSecurite(.Tab_Menu(nNumMenu)) Then
					sBool = sINI_GetChaineFichierIni(vsUtilisateur, "OPT_" & nNumMenu, sFicMenuIni)
                    If sBool = "1" Or sBool = "0" Then
                        .Tab_DroitMenu(nNumMenu) = CBool(sBool)
                    End If
                End If
			Next 
		End With
		
		bINI_Init_DroitSurMenu = True
		Exit Function
Erreur: 
		MSG_AfficheErreur(giERR_DROIT_SUR_MENU, "Erreur pendant initialisation des droits sur menu : " & Err.Number & "=>", Err.Description)
		
	End Function

    ' Fonction qui récupère la partie gauche du couple UID/PWD
    ' Soit l'utilisateur
    Public Function sRecupereUtilisateurIni(ByRef vsuser As String) As String

        Dim nIndex As Short
        sRecupereUtilisateurIni = ""

        nIndex = InStr(1, vsuser, "/")

        If nIndex > 0 Then
            sRecupereUtilisateurIni = Left(vsuser, nIndex - 1)
        End If

    End Function

    ' Fonction qui récupère la partie droite du couple UID/PWD
    ' Soit le mot de passe
    Public Function sRecupereMotDePasseIni(ByRef vsuser As String) As String

        Dim nIndex As Short
        sRecupereMotDePasseIni = ""

        nIndex = InStr(1, vsuser, "/")

        If nIndex > 0 Then
            sRecupereMotDePasseIni = Right(vsuser, Len(vsuser) - nIndex)
        End If

    End Function

End Module