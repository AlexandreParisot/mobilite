Option Strict Off
Option Explicit On
Module modLogDansFichiers
	'PREFIXE affect� au module = "LDF"
	
	Public Const giPROC_bInitMenu As Short = 1
	Public Const giPROC_bInitParametrage As Short = 2
	Public Const giPROC_bRecupereConfiguration As Short = 3
	Public Const giPROC_bLogInitialisationPDT As Short = 4
	Public Const giPROC_EcranPrincipal As Short = 5
	Public Const giPROC_bInitCodeBarrePDT As Short = 6


    'Ecriture du Log d'initialisation du PDT
    'Avec �criture de l'adresse IP ...
    Public Function bLDF_LogInitialisationPDT() As Boolean
		On Error GoTo Erreur
		Dim iFile As Short
		Dim sFichier As String
		Dim nIndex As Short
		
		
		sFichier = My.Application.Info.DirectoryPath & gCST_sREPERTOIRE_LOG & "PDT" & Mid(go_TRM.TerminalID, Len(go_TRM.TerminalID) - 2, 3)
		
		With gTab_General
			.sFichierLog = sFichier & ".log"
			.sFichierErr = sFichier & "_Err.log"
			
			If gTab_Configuration.bLog Then
				iFile = FreeFile
				
				FileOpen(iFile, .sFichierLog, OpenMode.Append)
				With go_TRM
					
					PrintLine(iFile, "")
					PrintLine(iFile, "NOUVELLE CONNEXION ****************************************************")
					PrintLine(iFile, "-----------------------------------------------------------------------")
					PrintLine(iFile, "Connexion le " & Today & " � " & TimeOfDay)
					PrintLine(iFile, "Application : " & My.Application.Info.Title & " - Version " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Revision)
					PrintLine(iFile, "-----------------------------------------------------------------------")
					PrintLine(iFile, "")
					PrintLine(iFile, "PDT_IP     : " & .TerminalID)
					PrintLine(iFile, "BATTERIE P : " & .MainBattery)
					PrintLine(iFile, "BATTERIE L : " & .LithiumBattery)
					PrintLine(iFile, "DISK       : " & .DiskSpace)
					PrintLine(iFile, "RAM        : " & .Memory)
					PrintLine(iFile, "TERMINAL   : " & .TerminalType)
					PrintLine(iFile, "LIGNES     : " & .TerminalHeight)
					PrintLine(iFile, "COLONNES   : " & .TerminalWidth)
					PrintLine(iFile, "WS_VERSION : " & .WireLessVersion)
					PrintLine(iFile, "PING       : " & .Ping())
					PrintLine(iFile, "")
					PrintLine(iFile, "Param�trage...: ")

                    With gTab_Configuration
                        PrintLine(iFile, "IP M3BE           : " & .sIP)
                        PrintLine(iFile, "Port d'�coute     : " & .sPort)
                        PrintLine(iFile, "Domaine           : " & .sDomaine)
                        PrintLine(iFile, "D�lai Msg         : " & .iDelaiMessage)
                        PrintLine(iFile, "Time Wait Connect : " & .lTimeWait & " ms")
                        PrintLine(iFile, "Soci�t�           : " & .sSociete)
                        PrintLine(iFile, "Division          : " & .sDivision)
                        PrintLine(iFile, "Depot             : " & .sDepot)
                        PrintLine(iFile, "Code barre        : " & .sCodeBar)
                        PrintLine(iFile, "Utilisateurs      : ")
                        For nIndex = 1 To UBound(.sProfil)
                            PrintLine(iFile, "                    " & nIndex & ":" & .sProfil(nIndex).sUtilisateur)
                        Next
                    End With
                End With
				FileClose(iFile)
			End If
		End With
		
		bLDF_LogInitialisationPDT = True
		
		Exit Function
Erreur: 
		LDF_LogErreurApplication(giPROC_bLogInitialisationPDT)
		
	End Function
	
	'Ecriture dans fichier log d'une erreur retourn�e par l'application
	Public Sub LDF_LogErreurApplication(ByRef viProcSource As Short, Optional ByRef vbErrSystem As Boolean = True)
		Dim iFile As Short
		Dim sMessage As String
		Dim sTitre As String
		
		sTitre = Today & " - " & TimeOfDay & " : "
		
		Select Case viProcSource
			
			Case giPROC_bInitMenu
				If vbErrSystem Then
					sMessage = "Erreur pendant l'initialisation du menu g�n�ral"
				Else
					sMessage = "Le menu g�n�ral ne s'est pas correctement initialis�"
				End If
				
			Case giPROC_bInitParametrage
				If vbErrSystem Then
					sMessage = "Erreur pendant l'initialisation du param�trage "
				Else
					sMessage = "Le param�trage n'est pas complet. L'une des clefs dans le fichier ( " & gCST_sFICHIER_INI & " ) n'est pas ou mal renseign�e. "
				End If
				
			Case giPROC_bRecupereConfiguration
				sMessage = "Probl�me de connexion avec le PDT.La lecture de la configuration du PDT a �chou�e."
				
			Case giPROC_bLogInitialisationPDT
				sMessage = "Erreur pendant l'�criture du fichier Log d'initialisation..."
				
			Case giPROC_bInitCodeBarrePDT
				sMessage = "Erreur pendant l'initialisation du code barre sur le PDT"

            Case Else
                sMessage = "Erreur inconnue...L'application va se terminer."
				
		End Select
		
		
		
		
		If Err.Number <> 0 Then
			sMessage = sMessage & Chr(13) & Err.Number & ":" & Err.Description
		End If
		
		sMessage = sTitre & sMessage
		
		If gTab_Configuration.bLog Then
			iFile = FreeFile
			FileOpen(iFile, gTab_General.sFichierErr, OpenMode.Append)
			
			PrintLine(iFile, "--------------------------------------------------------------")
			PrintLine(iFile, gTab_General.sPDT)
			PrintLine(iFile, sMessage)
			
			FileClose(iFile)
		End If
	End Sub

    'Ecrit l'erreur qui est apparu � l'utilisateur
    'Dans un fichier .log
    Public Sub LDF_AfficheErreurDansLog(ByRef vsErreur As String, ByRef vsArg1 As String, ByRef vsArg2 As String)
		Dim iFile As Short
		
		If gTab_Configuration.bLog Then
			
			iFile = FreeFile
			
			FileOpen(iFile, gTab_General.sFichierErr, OpenMode.Append)
			
			PrintLine(iFile, "--------------------------------------------------------------")
			PrintLine(iFile, "**ERREUR : Le " & Today & " � " & TimeOfDay & " par " & gTab_General.sPDT)
			PrintLine(iFile, vsErreur)
			PrintLine(iFile, vsArg1)
			PrintLine(iFile, vsArg2)
			
			FileClose(iFile)
			
		End If
		
	End Sub
	
	Public Sub LDF_LogPourTrace(ByRef vsMessage As String)
		Dim iFile As Short
		
		If gTab_Configuration.bLog Then
			iFile = FreeFile
			FileOpen(iFile, gTab_General.sFichierLog, OpenMode.Append)
			
			PrintLine(iFile, "--------------------------------------------------------------")
			PrintLine(iFile, "Le " & Today & " � " & TimeOfDay)
			PrintLine(iFile, gTab_General.sPDT)
			PrintLine(iFile, vsMessage)
			
			FileClose(iFile)
		End If
	End Sub

    'Recherche de l'existence d'un r�pertoire
    'S'il n'existe pas, le cr�e
    Public Function LDF_bRechercheRepertoire(ByRef vsRepertoire As String) As Boolean
		On Error GoTo Erreur
		Dim sRepInitial As String
		
		sRepInitial = My.Application.Info.DirectoryPath & "\" & vsRepertoire
        If Dir(sRepInitial, FileAttribute.Directory) = "" Then
            MkDir(sRepInitial)
        End If
		LDF_bRechercheRepertoire = True
		
		Exit Function
Erreur: 
		MSG_AfficheErreur(giERR_RECHERCHE_FICHIER, "Erreur :" & Err.Number, "(" & sRepInitial & ") " & Err.Description)
	End Function
	
	Public Function LDF_bSupprimeFichier(ByRef vsRepertoire As String, ByRef vsFichier As String) As Boolean
		On Error GoTo Erreur
		Dim sFichier As String
		
		sFichier = My.Application.Info.DirectoryPath & vsRepertoire & "\" & gTab_General.sPDT & "_" & vsFichier
		
        If Dir(sFichier, FileAttribute.Normal) <> "" Then
            Kill(sFichier)
            LDF_bSupprimeFichier = True
        End If
		Exit Function
Erreur: 
		MSG_AfficheErreur(giERR_RECHERCHE_FICHIER, "Erreur :" & Err.Number, "(" & sFichier & ") " & Err.Description)
		
	End Function
	
	
	Public Function LDF_bVerifieExistenceFichier(ByRef vsRepertoire As String, ByRef vsFichier As String) As Boolean
		On Error GoTo Erreur
		Dim sFichier As String
		
		sFichier = My.Application.Info.DirectoryPath & vsRepertoire & "\" & gTab_General.sPDT & "_" & vsFichier
		
        If Dir(sFichier, FileAttribute.Normal) <> "" Then
            LDF_bVerifieExistenceFichier = True
        End If
		Exit Function
Erreur: 
		MSG_AfficheErreur(giERR_RECHERCHE_FICHIER, "Erreur :" & Err.Number, "(" & sFichier & ") " & Err.Description)
		
	End Function
End Module