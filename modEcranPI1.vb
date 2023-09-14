Option Strict Off
Option Explicit On
Module modEcranPI1
    Dim msTitre As String

    'Entrée de l'option TID - Transfert inter-Dépôt
    Public Sub EcranPI1(ByRef vsTitre As String)
        Dim sFicIni As String = ""
        Dim bFinSaisie As Boolean = False
        Dim sIndexDeLivraison As String = ""
        Dim sNumLP As String = ""
        Dim sNumLigneLP As String = ""
        Dim sNumOD As String = ""
        Dim sDepotDebut As String = ""
        Dim sDepotFin As String = ""
        Dim sEmplacementArticleDepot As String = ""
        Dim sEmplacementParDefaut As String = ""
        Dim sEAN128Saisi As String = ""
        Dim sArticle As String = ""
        Dim sArticleLibelle As String = ""
        Dim sArticleType As String = ""
        Dim sLot As String = ""
        Dim sQuantite As String = ""
        Dim sQuantiteRestanteAPreparer As String = ""
        Dim sNumOrdre As String = ""
        Dim bSuiteTraitementOK As Boolean = False
        Dim sStatutIDStock As String = ""
        Dim sQuantiteAffectee As String = ""
        Dim sDatePeremption As String = ""

        Dim sPID As String = ""

        msTitre = vsTitre
        sFicIni = My.Application.Info.DirectoryPath & "\" & gCST_sFICHIER_INI

        If PI1_bSaisiePID(sPID, sQuantite, bFinSaisie) Then
            'If (API_TID_bValideLigneLP(sIndexDeLivraison, sNumLP, sDepotDebut, sEmplacementParDefaut, sArticle, sLot, sQuantite, sNumLigneLP, sQuantiteRestanteAPreparer)) Then

            'End If
        End If

    End Sub

    'Saisie de la PID
    Private Function PI1_bSaisiePID(ByRef vsPID As String, ByRef vsQuantite As String, ByRef vbFinSaisie As Boolean) As Boolean

        Dim nPosSeparateur As Short = 999

        Dim sScan As String = ""
        Dim iRes As Integer = 0
        Dim vsLibelle As String = ""
        Dim vsStock As String = ""
        Dim vsSITE As String = ""
        Dim vsSUNM As String = ""

        PI1_bSaisiePID = False
        vbFinSaisie = False
        vsPID = ""

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not PI1_bSaisiePID And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            'TID_AffichageSaisieLP()
            PI1_AffichageSaisiePID("", "", "", "", "")

            'Demande de saisie
            sScan = go_IO.RFInput("", 15, CHR_nCentrer(15), 3, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else
                        LDF_LogPourTrace("CAB PID LU ===== : " & Trim(sScan))

                        vsPID = Trim(sScan)

                        API_PI1_GetInfoPID(vsPID, vsLibelle, vsStock, vsSITE, vsSUNM)

                        PI1_AffichageSaisiePID(vsPID, vsLibelle, vsStock, vsSITE, vsSUNM)

                        'Demande de saisie
                        sScan = go_IO.RFInput("", 9, CHR_nCentrer(15), 11, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
                        iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
                        If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                            If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                                If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                                    vbFinSaisie = True
                                Else

                                    If sScan.Trim() <> vsPID.Trim And sScan.Trim() <> "" Then

                                        LDF_LogPourTrace("STOCK PID LU ===== : " & Trim(sScan))

                                        vsQuantite = Trim(sScan)

                                        API_PI1_bAjustementDuStock("MAI", "PIECES DET", vsPID, vsQuantite)

                                    End If

                                End If
                            Else
                                vbFinSaisie = True
                            End If
                        Else
                            WRL_GestionErreurPDT(iRes, "PI1_bSaisiePID")
                        End If

                    End If
                Else
                    vbFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "PI1_bSaisiePID")
            End If

        End While

    End Function

    'Affichage du titre pour saisie de l'index de livraison et du N° de LP
    Private Sub PI1_AffichageSaisiePID(ByRef vsPID As String, ByRef vsLibelle As String, ByRef vsStock As String, ByRef vsSITE As String, ByRef vsSUNM As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" INVENTAIRE PID ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 2, "   > " + "Scan PID", WirelessStudioOle.RFIOConstants.WLNORMAL + WirelessStudioOle.RFIOConstants.WFGCOLORBLUE + WirelessStudioOle.RFIOConstants.WFGCOLORLIGHT)
            .RFPrint(0, 3, "     " + vsPID, WirelessStudioOle.RFIOConstants.WLNORMAL + WirelessStudioOle.RFIOConstants.WFGCOLORLIGHT)
            .RFPrint(0, 4, "     " + vsLibelle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 5, "     " + vsSITE, WirelessStudioOle.RFIOConstants.WLNORMAL + WirelessStudioOle.RFIOConstants.WFGCOLORRED + WirelessStudioOle.RFIOConstants.WFGCOLORLIGHT)
            .RFPrint(0, 6, "     " + vsSUNM, WirelessStudioOle.RFIOConstants.WLNORMAL + WirelessStudioOle.RFIOConstants.WFGCOLORRED + WirelessStudioOle.RFIOConstants.WFGCOLORLIGHT)
            .RFPrint(0, 8, "   > " + "Stock actuel", WirelessStudioOle.RFIOConstants.WLNORMAL + WirelessStudioOle.RFIOConstants.WFGCOLORBLUE + WirelessStudioOle.RFIOConstants.WFGCOLORLIGHT)
            .RFPrint(0, 9, "     " + vsStock, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 10, "   > " + "Nouveau Stock", WirelessStudioOle.RFIOConstants.WLNORMAL + WirelessStudioOle.RFIOConstants.WFGCOLORBLUE + WirelessStudioOle.RFIOConstants.WFGCOLORLIGHT)
        End With
    End Sub

    'API=MMS310MI
    'Fonction=Update
    Public Function API_PI1_bAjustementDuStock(ByVal vsDepot As String, ByVal vsEmplacement As String, ByVal vsArticle As String, ByVal vsQuantite As String) As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""
        Dim nQuantiteEnStock As Long = 0
        Dim nQuantite As Long = 0
        Dim nNouvelleQuantite As Long = 0

        API_PI1_bAjustementDuStock = False

        ''Calcul de la nouvelle quantité en stock sur Emplacement/Article/Lot
        'If (IsNumeric(CHR_TransformeSeparateurPourNumerique(Trim(vsQuantite)))) Then
        '    nQuantite = CHR_TransformeSeparateurPourNumerique(Trim(vsQuantite))
        'End If
        'If (IsNumeric(CHR_TransformeSeparateurPourNumerique(Trim(vsNouvelleQuantite)))) Then
        '    nNouvelleQuantite = CHR_TransformeSeparateurPourNumerique(Trim(vsNouvelleQuantite))
        'End If
        'If (IsNumeric(CHR_TransformeSeparateurPourNumerique(Trim(vsQuantiteEnStock)))) Then
        '    nQuantiteEnStock = CHR_TransformeSeparateurPourNumerique(Trim(vsQuantiteEnStock))
        'End If

        'If (nNouvelleQuantite = nQuantite) Then
        '    API_EMS_bAjustementDuStock = True
        '    Exit Function
        'Else
        '    nQuantiteEnStock = nQuantiteEnStock - (nQuantite - nNouvelleQuantite)
        'End If



        If API_bConnexionAPI("MMS310MI") Then

            'Construction de la fonction et de ses paramètres pour l'appel API
            sParam = CHR_sAjoutEspace("Update", 15) &
                     CHR_sAjoutEspace(gTab_Configuration.sSociete, 3) &
                     CHR_sAjoutEspace(vsDepot, 3) &
                     CHR_sAjoutEspace(vsArticle, 15) &
                     CHR_sAjoutEspace(vsEmplacement, 10) &
                     CHR_sAjoutEspace("", 20) &
                     CHR_sAjoutEspace("", 30) &
                     CHR_sAjoutEspace(vsQuantite, 11) &
                     CHR_sAjoutEspace("", 26) &
                     CHR_sAjoutEspace(DateTime.Now.Date.ToString("yyyyMMdd"), 8)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then
                    API_PI1_bAjustementDuStock = True
                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

    'API=MMS200MI,PPS040MI,CRS620MI
    'Fonction=GetSumWhsBal,GetItmWhsBasic,GetItemSupplier
    Public Function API_PI1_GetInfoPID(ByVal vsPID As String, ByRef vsLibelle As String, ByRef vsStock As String, ByRef vsSITE As String, ByRef vsSUNM As String) As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""
        Dim vsSUNO As String = ""
        Dim nQuantiteEnStock As Long = 0
        Dim nQuantite As Long = 0
        Dim nNouvelleQuantite As Long = 0

        API_PI1_GetInfoPID = False

        If API_bConnexionAPI("MMS200MI") Then

            'Construction de la fonction et de ses paramètres pour l'appel API
            sParam = CHR_sAjoutEspace("GetSumWhsBal", 15) &
                     CHR_sAjoutEspace(gTab_Configuration.sSociete, 3) &
                     CHR_sAjoutEspace("MAI", 3) &
                     CHR_sAjoutEspace(vsPID, 30)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then

                    vsLibelle = Trim(Mid(sResultat, 1 + 3 + 3 + 30 + 15 + 36 + 3 + 10, 30))
                    vsStock = Trim(Mid(sResultat, 1 + 3 + 3 + 30 + 15 + 36 + 30 + 2 + 17 + 3 + 3 + 10 + 17 + 3 + 10 + 3 + 10, 17))

                    API_PI1_GetInfoPID = True

                Else

                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)

                End If
            End If

        End If

        If API_bConnexionAPI("MMS200MI") Then

            'Construction de la fonction et de ses paramètres pour l'appel API
            sParam = CHR_sAjoutEspace("GetItmWhsBasic", 15) &
                     CHR_sAjoutEspace(gTab_Configuration.sSociete, 3) &
                     CHR_sAjoutEspace("MAI", 3) &
                     CHR_sAjoutEspace(vsPID, 30)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then

                    '0140219858
                    vsSUNO = Trim(Mid(sResultat, 1 + 3 + 3 + 15 + 30 + 36 + 1 + 1 + 1 + 1 + 30 + 3 + 10 + 10 + 3 + 12, 10))

                    API_PI1_GetInfoPID = True

                Else

                    MSG_AfficheErreur(giERR_PROCEDURE_API, "1." + sResultat)

                End If
            End If

        End If

        If API_bConnexionAPI("PPS040MI") Then

            'Construction de la fonction et de ses paramètres pour l'appel API
            sParam = CHR_sAjoutEspace("GetItemSupplier", 15) &
                     CHR_sAjoutEspace(vsPID, 15) &
                     CHR_sAjoutEspace("", 3) &
                     CHR_sAjoutEspace("", 20) &
                     CHR_sAjoutEspace(vsSUNO, 10)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then

                    vsSITE = Trim(Mid(sResultat, 1 + 3 + 15 + 3 + 20 + 10 + 1 + 3 + 12, 30))

                    API_PI1_GetInfoPID = True

                Else

                    MSG_AfficheErreur(giERR_PROCEDURE_API, "2." + sResultat)

                End If
            End If

        End If

        If API_bConnexionAPI("CRS620MI") Then

            'Construction de la fonction et de ses paramètres pour l'appel API
            sParam = CHR_sAjoutEspace("GetBasicData", 15) &
                     CHR_sAjoutEspace(gTab_Configuration.sSociete, 3) &
                     CHR_sAjoutEspace(vsSUNO, 10)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then

                    vsSUNM = Trim(Mid(sResultat, 1 + 3 + 10 + 1 + 3 + 12, 36))

                    API_PI1_GetInfoPID = True

                Else

                    MSG_AfficheErreur(giERR_PROCEDURE_API, "3." + sResultat)

                End If
            End If

        End If


    End Function

End Module