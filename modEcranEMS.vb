Option Strict Off
Option Explicit On
Module modEcranEMS
    Dim msTitre As String


    'Entrée de l'option EMS - Edition d'une étiquette palette soit en lisant un EAN128 soit en saisissant toutes les informations utiles à la constitution du code à barre correspondant.
    'La quantité saisie va engendrer un ajustement de stock par rapport à la quantité initiale sur la palette.
    ' Le code à barre contient : N° d'ordre (AO ou OF) non obligatoire, code article MMS001, N° de lot, Quantité sur la palette
    Public Sub EcranEMS(ByRef vsTitre As String)
        Dim bFinSaisie As Boolean = False
        Dim sDepot As String = ""
        Dim sDepotLibelle As String = ""
        Dim sEmplacement As String = ""
        Dim sEAN128Saisi As String = ""
        Dim sNumOrdre As String = ""
        Dim sArticle As String = ""
        Dim sArticleLibelle As String = ""
        Dim sLot As String = ""
        Dim sQuantite As String = ""
        Dim sNouvelleQuantite As String = ""
        Dim sQuantiteEnStock As String = ""
        Dim sCodeMotif As String = ""
        Dim sCodeMotifLibelle As String = ""

        msTitre = vsTitre

        If EMS_bSaisieDepot(sDepot, sDepotLibelle) Then

            If EMS_bSaisieEmplacement(sDepot, sEmplacement) Then

                While Not bFinSaisie

                    sEAN128Saisi = ""
                    sArticle = ""
                    sLot = ""
                    sQuantite = ""
                    EMS_bChoixScanOuSaisieEAN128(sDepot, sEmplacement, sEAN128Saisi, bFinSaisie)

                    If (Not bFinSaisie) Then

                        ' ====================
                        'Le CAB a été scanné ou saisi
                        If (sEAN128Saisi <> "") Then
                            If (CHR_bRecupInfoEAN128(sEAN128Saisi, sNumOrdre, sArticle, sLot, sQuantite)) Then

                                If (API_EMS_bControleArticleLot(sNumOrdre, sArticle, sLot)) Then

                                    If (API_EMS_bControleQuantiteArticleLot(sDepot, sArticle, sLot, sQuantite, sEmplacement, sQuantiteEnStock)) Then

                                        If (EMS_bSaisieCodeMotif(sDepot, sEmplacement, sArticle, sArticleLibelle, sLot, sNumOrdre, sQuantite, sCodeMotif, sCodeMotifLibelle, bFinSaisie)) Then

                                            If (EMS_bSaisieNouvelleQuantite(sDepot, sEmplacement, sArticle, sArticleLibelle, sLot, sNumOrdre, sQuantite, sNouvelleQuantite, bFinSaisie)) Then

                                                If (API_EMS_bAjustementDuStock(sDepot, sEmplacement, sArticle, sLot, sQuantite, sNouvelleQuantite, sQuantiteEnStock, sCodeMotif)) Then
                                                    API_EMS_bEditionEtiquettePalette(sDepot, sEmplacement, sNumOrdre, sArticle, sLot, sNouvelleQuantite)
                                                End If

                                            End If

                                        End If

                                    End If

                                End If

                            End If

                        Else
                            ' ====================
                            'Les éléments du CAB doivent être saisis Article + Lot + Quantité
                            If (EMS_bSaisieArticle(sDepot, sEmplacement, sArticle, sArticleLibelle, bFinSaisie)) Then

                                If (EMS_bSaisieLot(sDepot, sEmplacement, sArticle, sArticleLibelle, sLot, sNumOrdre, bFinSaisie)) Then

                                    If (EMS_bSaisieQuantite(sDepot, sEmplacement, sArticle, sArticleLibelle, sLot, sNumOrdre, sQuantite, sQuantiteEnStock, bFinSaisie)) Then

                                        If (EMS_bSaisieCodeMotif(sDepot, sEmplacement, sArticle, sArticleLibelle, sLot, sNumOrdre, sQuantite, sCodeMotif, sCodeMotifLibelle, bFinSaisie)) Then

                                            If (EMS_bSaisieNouvelleQuantite(sDepot, sEmplacement, sArticle, sArticleLibelle, sLot, sNumOrdre, sQuantite, sNouvelleQuantite, bFinSaisie)) Then

                                                If (API_EMS_bAjustementDuStock(sDepot, sEmplacement, sArticle, sLot, sQuantite, sNouvelleQuantite, sQuantiteEnStock, sCodeMotif)) Then
                                                    API_EMS_bEditionEtiquettePalette(sDepot, sEmplacement, sNumOrdre, sArticle, sLot, sNouvelleQuantite)
                                                End If

                                            End If

                                        End If

                                    End If

                                End If

                            End If


                        End If
                    End If

                End While

            End If

        End If

    End Sub

    'Saisie du dépôt
    Private Function EMS_bSaisieDepot(ByRef vsDepot As String, ByRef vsDepotLibelle As String) As Boolean

        EMS_bSaisieDepot = False
        Dim bFinSaisie As Boolean = False
        Dim sScan As String = ""
        Dim iRes As Integer = 0

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not EMS_bSaisieDepot And Not gbErreurCommunication And Not bFinSaisie

            'Affichage
            EMS_AffichageSaisieDEPOT(vsDepot)

            'Demande de saisie
            sScan = go_IO.RFInput(gTab_Configuration.sDepot, 3, CHR_nCentrer(3), 3, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        bFinSaisie = True
                    Else

                        vsDepot = Trim(sScan)
                        If (API_EEP_bRechercheDepot(vsDepot, vsDepotLibelle)) Then
                            EMS_bSaisieDepot = True
                        End If

                    End If
                Else
                    bFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "EMS_bSaisieDepot")
            End If

        End While

    End Function

    'Saisie de l'emplacement
    Private Function EMS_bSaisieEmplacement(ByVal vsDepot As String, ByRef vsEmplacement As String) As Boolean

        EMS_bSaisieEmplacement = False
        Dim bFinSaisie As Boolean = False
        Dim sScan As String = ""
        Dim iRes As Integer = 0

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not EMS_bSaisieEmplacement And Not gbErreurCommunication And Not bFinSaisie

            'Affichage
            EMS_AffichageSaisieEMPLACEMENT(vsDepot, vsEmplacement)

            'Demande de saisie
            sScan = go_IO.RFInput("", 10, CHR_nCentrer(10), 5, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        bFinSaisie = True
                    Else

                        vsEmplacement = Trim(sScan)
                        If (API_EMS_bRechercheEmplacement(vsDepot, vsEmplacement)) Then
                            EMS_bSaisieEmplacement = True
                        End If

                    End If
                Else
                    bFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "EMS_bSaisieEmplacement")
            End If

        End While

    End Function

    'Choix, soit Scan ou Saisie du code EAN128
    Private Function EMS_bChoixScanOuSaisieEAN128(ByVal vsDepot As String, ByVal vsEmplacement As String, ByRef vsEAN128Saisi As String, ByRef vbFinSaisie As Boolean) As Boolean

        EMS_bChoixScanOuSaisieEAN128 = False
        Dim sScan As String = ""
        Dim iRes As Integer = 0
        vsEAN128Saisi = ""

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not EMS_bChoixScanOuSaisieEAN128 And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            EMS_AffichageChoix(vsDepot, vsEmplacement)

            'Demande de saisie
            sScan = go_IO.RFInput("", 70, CHR_nCentrer(70), 6, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT_SAISIE, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else
                        If sScan = gCST_TOUCHE_SAISIE_F5 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                            EMS_bChoixScanOuSaisieEAN128 = True
                        Else
                            vsEAN128Saisi = sScan
                            EMS_bChoixScanOuSaisieEAN128 = True
                        End If
                    End If
                Else
                    vbFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "EMS_bChoixSsanOuSaisieEAN128")
            End If
        End While

    End Function

    'Saisie de l'article
    Private Function EMS_bSaisieArticle(ByVal vsDepot As String, ByVal vsEmplacement As String, ByRef vsArticle As String, ByRef vsArticleLibelle As String, ByRef vbFinSaisie As Boolean) As Boolean

        EMS_bSaisieArticle = False
        Dim sScan As String = ""
        Dim iRes As Integer = 0

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not EMS_bSaisieArticle And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            EMS_AffichageSaisieARTICLE(vsDepot, vsEmplacement)

            'Demande de saisie
            sScan = go_IO.RFInput("", 15, CHR_nCentrer(15), 6, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else

                        vsArticle = Trim(sScan)
                        If (API_EMS_bRechercheArticle(vsArticle, vsArticleLibelle)) Then
                            EMS_bSaisieArticle = True
                        End If

                    End If
                Else
                    vbFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "EMS_bSaisieArticle")
            End If

        End While

    End Function

    'Saisie du lot
    Private Function EMS_bSaisieLot(ByVal vsDepot As String, ByVal vsEmplacement As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByRef vsLot As String, ByRef vsNumOrdre As String, ByRef vbFinSaisie As Boolean) As Boolean

        EMS_bSaisieLot = False
        Dim sScan As String = ""
        Dim iRes As Integer = 0
        vsLot = ""
        vsNumOrdre = ""

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not EMS_bSaisieLot And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            EMS_AffichageSaisieLOT(vsDepot, vsEmplacement, vsArticle, vsArticleLibelle)

            'Demande de saisie
            sScan = go_IO.RFInput("", 20, CHR_nCentrer(20), 8, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else

                        vsLot = Trim(sScan)
                        If (API_EMS_bControleArticleLot(vsNumOrdre, vsArticle, vsLot)) Then
                            EMS_bSaisieLot = True
                        End If

                    End If
                Else
                    vbFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "EMS_bSaisieLot")
            End If

        End While

    End Function

    'Saisie de la quantité
    Private Function EMS_bSaisieQuantite(ByVal vsDepot As String, ByVal vsEmplacement As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByVal vsLot As String, ByVal vsNumOrdre As String,
                                         ByRef vsQuantite As String, ByRef vsQuantiteEnStock As String, ByRef vbFinSaisie As Boolean) As Boolean

        Dim sScan As String = ""
        Dim iRes As Integer = 0
        Dim nQuantite As Long = 0

        EMS_bSaisieQuantite = False
        vsQuantite = ""
        vsQuantiteEnStock = ""

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not EMS_bSaisieQuantite And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            EMS_AffichageSaisieQUANTITE(vsDepot, vsEmplacement, vsArticle, vsArticleLibelle, vsLot, vsNumOrdre)

            'Demande de saisie
            sScan = go_IO.RFInput("", 10, CHR_nCentrer(10), 10, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else
                        If (IsNumeric(CHR_TransformeSeparateurPourNumerique(Trim(sScan)))) Then
                            nQuantite = (CHR_TransformeSeparateurPourNumerique(Trim(sScan)))
                            If (nQuantite <> 0) Then
                                vsQuantite = Trim(sScan)
                                If (API_EMS_bControleQuantiteArticleLot(vsDepot, vsArticle, vsLot, vsQuantite, vsEmplacement, vsQuantiteEnStock)) Then
                                    EMS_bSaisieQuantite = True
                                End If
                            Else
                                MSG_AfficheErreur(giERR_QUANTITE_NULL)
                            End If

                        Else
                            MSG_AfficheErreur(giERR_FORMAT_NUMERIC)
                        End If

                    End If
                Else
                    vbFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "EMS_bSaisieQuantite")
            End If

        End While

    End Function

    'Saisie du code motif
    Private Function EMS_bSaisieCodeMotif(ByVal vsDepot As String, ByVal vsEmplacement As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByVal vsLot As String, ByVal vsNumOrdre As String,
                                          ByVal vsQuantite As String, ByRef vsCodeMotif As String, ByRef vsCodeMotifLibelle As String, ByRef vbFinSaisie As Boolean) As Boolean

        Dim sScan As String = ""
        Dim iRes As Integer = 0

        EMS_bSaisieCodeMotif = False
        vsCodeMotif = ""
        vsCodeMotifLibelle = ""

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not EMS_bSaisieCodeMotif And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            EMS_AffichageSaisieCodeMotif(vsDepot, vsEmplacement, vsArticle, vsArticleLibelle, vsLot, vsNumOrdre, vsQuantite)

            'Demande de saisie
            sScan = go_IO.RFInput("", 10, CHR_nCentrer(10), 11, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else
                        vsCodeMotif = Trim(sScan)
                        If (API_EMS_bControleCodeMotif(vsCodeMotif, vsCodeMotifLibelle)) Then
                            EMS_bSaisieCodeMotif = True
                        End If
                    End If
                Else
                    vbFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "EMS_bSaisieCodeMotif")
            End If

        End While

    End Function

    'Saisie de la nouvelle quantité
    Private Function EMS_bSaisieNouvelleQuantite(ByVal vsDepot As String, ByVal vsEmplacement As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByVal vsLot As String, ByVal vsNumOrdre As String,
                                                 ByVal vsQuantite As String, ByRef vsNouvelleQuantite As String, ByRef vbFinSaisie As Boolean) As Boolean

        Dim sScan As String = ""
        Dim iRes As Integer = 0
        Dim nQuantite As Long = 0
        Dim nNouvelleQuantite As Long = 0

        EMS_bSaisieNouvelleQuantite = False
        vsNouvelleQuantite = ""

        If (IsNumeric(CHR_TransformeSeparateurPourNumerique(Trim(vsQuantite)))) Then
            nQuantite = (CHR_TransformeSeparateurPourNumerique(Trim(vsQuantite)))
        End If

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not EMS_bSaisieNouvelleQuantite And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            EMS_AffichageSaisieNouvelleQUANTITE(vsDepot, vsEmplacement, vsArticle, vsArticleLibelle, vsLot, vsNumOrdre, vsQuantite)

            'Demande de saisie
            sScan = go_IO.RFInput(vsQuantite, 10, CHR_nCentrer(10), 11, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else
                        If (IsNumeric(CHR_TransformeSeparateurPourNumerique(Trim(sScan)))) Then
                            nNouvelleQuantite = (CHR_TransformeSeparateurPourNumerique(Trim(sScan)))
                            If (nNouvelleQuantite <> 0) Then

                                If (nNouvelleQuantite > nQuantite) Then
                                    MSG_AfficheErreur(giERR_QUANTITE_DOIT_ETRE_INFERIEURE)
                                Else
                                    vsNouvelleQuantite = Trim(sScan)
                                    If (API_EMS_bControleQuantiteArticleLot(vsDepot, vsArticle, vsLot, vsQuantite, vsEmplacement)) Then
                                        EMS_bSaisieNouvelleQuantite = True
                                    End If
                                End If

                            Else
                                MSG_AfficheErreur(giERR_QUANTITE_NULL)
                            End If

                        Else
                            MSG_AfficheErreur(giERR_FORMAT_NUMERIC)
                        End If

                    End If
                Else
                    vbFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "EMS_bSaisieNouvelleQuantite")
            End If

        End While

    End Function

    'Affichage du titre pour saisie du dépôt
    Private Sub EMS_AffichageSaisieDEPOT(ByVal vsDepot As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 2, CHR_sCentrer("DEPOT"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage du titre pour saisie de l'emplacement
    Private Sub EMS_AffichageSaisieEMPLACEMENT(ByVal vsDepot As String, ByVal vsEmplacement As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 2, "DEPOT : " & vsDepot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 4, CHR_sCentrer("EMPLACEMENT"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage de l'écran choix de la saisie (Scan CAB ou saisie manuelle)
    Private Sub EMS_AffichageChoix(ByVal vsDepot As String, ByVal vsEmplacement As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 2, "DEPOT : " & vsDepot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, "EMPL. : " & vsEmplacement, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 5, CHR_sCentrer("CAB PALETTE"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage du titre pour saisie de l'article
    Private Sub EMS_AffichageSaisieARTICLE(ByVal vsDepot As String, ByVal vsEmplacement As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 2, "DEPOT : " & vsDepot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, "EMPL. : " & vsEmplacement, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 5, CHR_sCentrer("ARTICLE"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage du titre pour saisie du lot
    Private Sub EMS_AffichageSaisieLOT(ByVal vsDepot As String, ByVal vsEmplacement As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 2, "DEPOT : " & vsDepot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, "EMPL. : " & vsEmplacement, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 4, "ART.  : " & vsArticle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 5, vsArticleLibelle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 7, CHR_sCentrer("LOT"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage du titre pour saisie de la quantité
    Private Sub EMS_AffichageSaisieQUANTITE(ByVal vsDepot As String, ByVal vsEmplacement As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByVal vsLot As String, ByVal vsNumOrdre As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 2, "DEPOT : " & vsDepot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, "EMPL. : " & vsEmplacement, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 4, "ART.  : " & vsArticle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 5, vsArticleLibelle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 6, "LOT   : " & vsLot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 7, "N° ORDRE : " & vsNumOrdre, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 9, CHR_sCentrer("QUANTITE"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage du titre pour saisie du code motif
    Private Sub EMS_AffichageSaisieCodeMotif(ByVal vsDepot As String, ByVal vsEmplacement As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByVal vsLot As String, ByVal vsNumOrdre As String,
                                             ByVal vsQuantite As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 2, "DEPOT : " & vsDepot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, "EMPL. : " & vsEmplacement, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 4, "ART.  : " & vsArticle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 5, vsArticleLibelle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 6, "LOT   : " & vsLot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 7, "N° ORDRE : " & vsNumOrdre, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 8, "QTE   : " & vsQuantite, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 10, CHR_sCentrer("CODE MOTIF"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage du titre pour saisie de la nouvelle quantité
    Private Sub EMS_AffichageSaisieNouvelleQUANTITE(ByVal vsDepot As String, ByVal vsEmplacement As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByVal vsLot As String, ByVal vsNumOrdre As String,
                                                    ByVal vsQuantite As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 2, "DEPOT : " & vsDepot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, "EMPL. : " & vsEmplacement, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 4, "ART.  : " & vsArticle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 5, vsArticleLibelle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 6, "LOT   : " & vsLot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 7, "N° ORDRE : " & vsNumOrdre, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 8, "QTE   : " & vsQuantite, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 10, CHR_sCentrer("NOUVELLE QUANTITE"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

End Module