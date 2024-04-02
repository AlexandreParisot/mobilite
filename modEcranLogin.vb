Option Strict Off
Option Explicit On
Module modEcranLogin

    'Identification utilisateur
    Public Function bELG_SaisieLogin() As Boolean

        Dim sScan As String = ""
        Dim iRes As Integer

        With gTab_Configuration

            If .sUtilisateur = "" And .sMotDePasse = "" Then

                While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not bELG_SaisieLogin And Not gbErreurCommunication

                    'Affichage
                    ELG_AffichageTitreLogin("")

                    'Demande de saisie de l'utilisateur
                    sScan = go_IO.RFInput("", 10, CHR_nCentrer(10), 3, gCST_sFICHIER_CODE_BARRE, WirelessStudioOle.RFIOConstants.WLNORMALKEYS, WirelessStudioOle.RFIOConstants.WLMAXLENGTH + WirelessStudioOle.RFIOConstants.WLNO_RETURN_FILL + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
                    iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
                    If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then

                        If sScan <> Chr(gCST_TOUCHE_ECHAP) Then

                            .sUtilisateur = Trim(sScan)

                            ELG_AffichageTitreLogin(sScan)

                            'Demande de saisie
                            sScan = go_IO.RFInput("", 14, CHR_nCentrer(14), 6, gCST_sFICHIER_CODE_BARRE, WirelessStudioOle.RFIOConstants.WLNORMALKEYS, WirelessStudioOle.RFIOConstants.WLMAXLENGTH + WirelessStudioOle.RFIOConstants.WLNO_RETURN_FILL + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
                            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
                            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then

                                    .sMotDePasse = Trim(sScan)

                                    If bELG_VerifieUtilisateur() Then

                                        bELG_SaisieLogin = True
                                    Else
                                        MSG_AfficheErreur(giERR_LOGIN_USER)
                                    End If
                                End If
                            Else
                                WRL_GestionErreurPDT(iRes, "bELG_SaisieLogin")
                            End If
                        End If
                    Else
                        WRL_GestionErreurPDT(iRes, "bELG_SaisieLogin")
                    End If
                End While
            Else
                bELG_SaisieLogin = True
            End If
        End With

    End Function


    'Verification que l'utilisateur et le mot de passe saisie correspondent au paramètrage
    Private Function bELG_VerifieUtilisateur() As Boolean
        Dim sParam As String = ""
        Dim sResultat As String = ""
        Dim nIndex As Short

        Try

            With gTab_Configuration
                For nIndex = 1 To UBound(.sProfil)
                    If .sProfil(nIndex).sUtilisateur = .sUtilisateur Then

                        If API_bConnexionAPI("GENERAL") Then
                            'Construction de la fonction et de ses paramètres pour l'appel API
                            sParam = CHR_sAjoutEspace("GetUserInfo", 15)

                            If API_bTraitementAPI(sParam, sResultat) Then
                                If Mid(sResultat, 1, 3) <> "NOK" Then
                                    ' Récupération du dépôt de l'utilisateur (MNS150)
                                    If (Mid(sResultat, 35, 3) <> "") Then
                                        .sDepot = Mid(sResultat, 35, 3)
                                    End If
                                    bELG_VerifieUtilisateur = True
                                Else
                                    bELG_VerifieUtilisateur = False
                                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                                End If
                            End If
                        End If

                        Exit For
                    End If
                Next
            End With

        Catch ex2 As TypeInitializationException
            LDF_AfficheErreurDansLog("2", "0", ex2.InnerException.ToString)
        Catch ex As Exception
            LDF_AfficheErreurDansLog("1", "0", ex.ToString)
        End Try


    End Function

    'Affichage du titre pour la saisie du login
    Private Sub ELG_AffichageTitreLogin(ByRef vsuser As String)
		With go_IO
			.RFPrint(0, 0, CHR_sCentrer(" IDENTIFICATION ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
			.RFPrint(0, 2, CHR_sCentrer("Utilisateur"), WirelessStudioOle.RFIOConstants.WLNORMAL)
			.RFPrint(0, 3, CHR_sCentrer(vsuser), WirelessStudioOle.RFIOConstants.WLNORMAL)
			.RFPrint(0, 5, CHR_sCentrer("Mot de passe"), WirelessStudioOle.RFIOConstants.WLNORMAL)
		End With
	End Sub


End Module