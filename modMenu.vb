Option Strict Off
Option Explicit On
Module modMenu

    'PREFIXE affecté au module = "MNU"


    'Variables liées au menu général
    'Leur valeur est renseignée, dans l'ajout des options
    Public giOPT_EEP_EDITION_ETIQUETTE_PALETTE As Short
    Public giOPT_FIN As Short

    'Fonction d'entrée pour l'initialisation des menus
    'Retourne True si tous les menus se sont correctement initialisés
    Public Function bMNU_InitMenu() As Boolean
		
		If bINI_Init_DroitSurMenu(gTab_Configuration.sUtilisateur) Then
            If bMNU_InitMenuGeneral() Then
                bMNU_InitMenu = True
            End If
        End If
		
	End Function

    'Fonction d'initialisation du menu général
    'Retourne True si Ok
    Public Function bMNU_InitMenuGeneral() As Boolean
		On Error GoTo Erreur
		
		With go_MNU
			'Menu Général
			If .ResetMenu() Then
				.AddTitleLine("MENU")
				
				MNU_AjouteOptionEnFonctionDesDroits()
				
				.SetCoordinates(0, 0, go_TRM.TerminalWidth, go_TRM.TerminalHeight)
				If .StoreMenu(gCST_sFICHIER_MNU_GENERAL) Then
					bMNU_InitMenuGeneral = True
				End If
			End If
		End With
		
		Exit Function
Erreur: 
		MSG_AfficheErreur(giERR_INITIALISATION_MENU, "Erreur pendant l'affichage du menu general : " & Err.Number & "=>", Err.Description)
	End Function

    'Procedure pour l'ajout des options au menu général
    'On utilise gTab_DroitSurMenu pour obtenir le droit
    'Si true : On ajoute l'option à go_MNU
    Private Sub MNU_AjouteOptionEnFonctionDesDroits()
		Dim nOption As Short
		Dim nMenu As Short
		
		With gTab_Menu
			
			For nMenu = 1 To .nNombreMenu
				If .Tab_DroitMenu(nMenu) Then
					nOption = nOption + 1
					
					ReDim Preserve .Tab_Option(nOption)
					.Tab_Option(nOption) = Mid(.Tab_Menu(nMenu), 1, 3)
					
					go_MNU.AddOption(nOption & "-" & Mid(.Tab_Menu(nMenu), 5))
				End If
			Next 
			
			nOption = nOption + 1
			If nOption > 1 Then
                go_MNU.AddOption(nOption & " Fin")
            Else
				go_MNU.AddOption(nOption & " Aucun menu")
			End If
			giOPT_FIN = nOption
			
		End With
		
	End Sub
End Module