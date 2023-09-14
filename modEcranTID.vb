Option Strict Off
Option Explicit On
Module modEcranTID
    Dim msTitre As String

    'Entrée de l'option TID - Transfert inter-Dépôt
    Public Sub EcranTID(ByRef vsTitre As String)
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

        msTitre = vsTitre
        sFicIni = My.Application.Info.DirectoryPath & "\" & gCST_sFICHIER_INI

        TID_bSaisieLP(sIndexDeLivraison, sNumLP, sNumOD, sDepotDebut, sDepotFin, bFinSaisie)

        While Not bFinSaisie

            sEAN128Saisi = ""
            sArticle = ""
            sLot = ""
            sQuantite = ""
            sEmplacementParDefaut = ""
            TID_bChoixScanOuSaisieEAN128(sIndexDeLivraison, sNumLP, sDepotDebut, sEmplacementParDefaut, sEAN128Saisi, bFinSaisie)

            If (Not bFinSaisie) Then

                ' ====================
                'Le CAB a été scanné ou saisi
                If (sEAN128Saisi <> "") Then
                    If (CHR_bRecupInfoEAN128(sEAN128Saisi, sNumOrdre, sArticle, sLot, sQuantite)) Then

                        If (API_TID_bRechercheArticle(sArticle, sArticleLibelle, sArticleType)) Then

                            If (API_TID_bControleArticleLot(sNumOrdre, sArticle, sLot)) Then

                                If (API_TID_bControleArticleDepot(sDepotDebut, sArticle, sEmplacementArticleDepot)) Then

                                    'On indique quel est l'emplacement de départ pour le transfert de marchandise
                                    sEmplacementParDefaut = sINI_GetChaineFichierIni(gCST_INI_SEC_TID, "EMPLACEMENT_DEPART_" + sArticleType, sFicIni)
                                    If (sEmplacementParDefaut = "") Then
                                        sEmplacementParDefaut = sEmplacementArticleDepot
                                    End If

                                    If (API_TID_bControleQuantiteArticleLot(sDepotDebut, sArticle, sLot, sQuantite, sEmplacementParDefaut, sStatutIDStock, sQuantiteAffectee, sDatePeremption)) Then

                                        bSuiteTraitementOK = True

                                    End If

                                End If

                            End If

                        End If

                    End If

                Else
                    ' ====================
                    'Les éléments du CAB doivent être saisis Article + Lot
                    sArticle = ""
                    sLot = ""
                    sQuantite = ""
                    If (TID_bSaisieArticle(sArticle, sArticleLibelle, sArticleType, bFinSaisie)) Then

                        If (API_TID_bControleArticleDepot(sDepotDebut, sArticle, sEmplacementArticleDepot)) Then

                            If (TID_bSaisieLot(sEmplacementParDefaut, sArticle, sArticleLibelle, sLot, sNumOrdre, bFinSaisie)) Then

                                'On indique quel est l'emplacement de départ pour le transfert de marchandise
                                sEmplacementParDefaut = sINI_GetChaineFichierIni(gCST_INI_SEC_TID, "EMPLACEMENT_DEPART_" + sArticleType, sFicIni)
                                If (sEmplacementParDefaut = "") Then
                                    sEmplacementParDefaut = sEmplacementArticleDepot
                                End If

                                If (TID_bSaisieQuantite(sDepotDebut, sEmplacementParDefaut, sArticle, sArticleLibelle, sLot, sNumOrdre, sQuantite, bFinSaisie)) Then

                                    bSuiteTraitementOK = True

                                End If

                            End If

                        End If

                    End If

                End If

            End If


            ' ====================
            'Suite du traitement
            If (Not bFinSaisie And bSuiteTraitementOK = True) Then
                'A ce stade on a recueilli les info Article/Lot (et la quantité si on a scannné le CAB) et on les a vérifié

                If (API_TID_bRechercheLigneLP_Article(sIndexDeLivraison, sNumLP, sArticle, sNumLigneLP, sQuantiteRestanteAPreparer)) Then

                    If (TID_bSaisieQuantitePreparee(sIndexDeLivraison, sNumLP, sDepotDebut, sEmplacementParDefaut, sArticle, sArticleLibelle, sLot, sNumOrdre, sQuantite, sQuantiteRestanteAPreparer, bFinSaisie)) Then

                        If (API_TID_bValideLigneLP(sIndexDeLivraison, sNumLP, sDepotDebut, sEmplacementParDefaut, sArticle, sLot, sQuantite, sNumLigneLP, sQuantiteRestanteAPreparer)) Then
                        End If

                    End If

                End If

            End If

        End While

    End Sub

    'Saisie de l'Index de livraison et du Numéro de LP
    'Le CAB est sous la forme Index de livraison + "/" + Numéro de LP
    Private Function TID_bSaisieLP(ByRef vsIndexDeLivraison As String, ByRef vsNumLP As String, ByRef vsNumOD As String, ByRef vsDepotDebut As String,
                                   ByRef vsDepotFin As String, ByRef vbFinSaisie As Boolean) As Boolean

        Dim nPosSeparateur As Short = 999

        Dim sScan As String = ""
        Dim iRes As Integer = 0

        TID_bSaisieLP = False
        vbFinSaisie = False
        vsIndexDeLivraison = ""
        vsNumLP = ""

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not TID_bSaisieLP And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            TID_AffichageSaisieLP()

            'Demande de saisie
            sScan = go_IO.RFInput("", 15, CHR_nCentrer(15), 3, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else
                        LDF_LogPourTrace("CAB INDEX/LP LU ===== : " & Trim(sScan))

                        'Recherche du "/" qui sépare les 2 valeures
                        nPosSeparateur = InStr(1, Trim(sScan), "/")

                        If (nPosSeparateur > 0) Then
                            'Recupération de l'index de livraison
                            vsIndexDeLivraison = Mid(Trim(sScan), 1, (nPosSeparateur - 1))
                            'Recupération du N° de LP
                            vsNumLP = Mid(Trim(sScan), nPosSeparateur + 1, Trim(sScan).Length)
                        End If

                        If (vsIndexDeLivraison <> "" And vsNumLP <> "") Then
                            If (API_TID_bRechercheLP(vsIndexDeLivraison, vsNumLP)) Then
                                If (API_TID_bRtvInfoLP(vsIndexDeLivraison, vsNumLP, vsNumOD, vsDepotDebut, vsDepotFin)) Then
                                    TID_bSaisieLP = True
                                End If
                            End If
                        Else
                            MSG_AfficheErreur(giERR_CAB_INVALIDE)
                        End If

                    End If
                Else
                    vbFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "TID_bSaisieLP")
            End If

        End While

    End Function

    'Choix, soit Scan ou Saisie du code EAN128
    Private Function TID_bChoixScanOuSaisieEAN128(ByVal vsIndexDeLivraison As String, ByVal vsNumLP As String, ByVal vsDepotDebut As String, ByVal vsEmplacement As String, ByRef vsEAN128Saisi As String, ByRef vbFinSaisie As Boolean) As Boolean

        Dim sScan As String = ""
        Dim iRes As Integer = 0
        Dim sArticle As String = ""
        Dim sArticleLibelle As String = ""
        Dim sQuantite As String = ""
        Dim sNumLigneLP As String = ""

        TID_bChoixScanOuSaisieEAN128 = False
        vsEAN128Saisi = ""


        API_TID_bRechercheLigneLP_A_Preparer(vsIndexDeLivraison, vsNumLP, sArticle, sArticleLibelle, sQuantite, sNumLigneLP, "")


        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not TID_bChoixScanOuSaisieEAN128 And Not gbErreurCommunication And Not vbFinSaisie

            If (sNumLigneLP = "") Then

                'Affichage
                TID_AffichageValidationLP(vsIndexDeLivraison, vsNumLP)

                'Demande de saisie
                sScan = go_IO.GetEventEx(gCST_sFICHIER_BOUTONS_OK_CLR_QUIT)
                iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
                If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                    If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                        If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                            vbFinSaisie = True
                        Else
                            API_TID_bValideLP(vsIndexDeLivraison, vsNumLP)
                            vbFinSaisie = True
                        End If
                    Else
                        vbFinSaisie = True
                    End If
                Else
                    WRL_GestionErreurPDT(iRes, "TID_bChoixSsanOuSaisieEAN128")
                End If

            Else

                'Affichage
                TID_AffichageChoix(vsIndexDeLivraison, vsNumLP, sArticle, sArticleLibelle, sQuantite)

                'Demande de saisie
                sScan = go_IO.RFInput("", 70, CHR_nCentrer(70), 7, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT_PRECEDENT_SUIVANT_SAISIE_FIN_LIGNE, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
                iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
                If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                    If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                        If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                            vbFinSaisie = True
                        Else
                            If sScan = gCST_TOUCHE_PRECEDENT_F1 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                                If (API_TID_bRechercheLigneLP_A_Preparer(vsIndexDeLivraison, vsNumLP, sArticle, sArticleLibelle, sQuantite, sNumLigneLP, "-") = False) Then
                                    API_TID_bRechercheLigneLP_A_Preparer(vsIndexDeLivraison, vsNumLP, sArticle, sArticleLibelle, sQuantite, sNumLigneLP, "")
                                End If
                            Else
                                If sScan = gCST_TOUCHE_SUIVANT_F2 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                                    If (API_TID_bRechercheLigneLP_A_Preparer(vsIndexDeLivraison, vsNumLP, sArticle, sArticleLibelle, sQuantite, sNumLigneLP, "+") = False) Then
                                        API_TID_bRechercheLigneLP_A_Preparer(vsIndexDeLivraison, vsNumLP, sArticle, sArticleLibelle, sQuantite, sNumLigneLP, "")
                                    End If
                                Else
                                    If sScan = gCST_TOUCHE_SAISIE_F5 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                                        TID_bChoixScanOuSaisieEAN128 = True
                                    Else
                                        If sScan = gCST_TOUCHE_FIN_LIGNE And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                                            API_TID_bValideLigneLP(vsIndexDeLivraison, vsNumLP, vsDepotDebut, vsEmplacement, "", "", "0", sNumLigneLP, sQuantite)
                                            sNumLigneLP = ""
                                            sArticle = ""
                                            sQuantite = ""
                                            API_TID_bRechercheLigneLP_A_Preparer(vsIndexDeLivraison, vsNumLP, sArticle, sArticleLibelle, sQuantite, sNumLigneLP, "")
                                        Else
                                            vsEAN128Saisi = sScan
                                            TID_bChoixScanOuSaisieEAN128 = True
                                        End If


                                    End If
                                End If
                            End If
                        End If
                    Else
                        vbFinSaisie = True
                    End If
                Else
                    WRL_GestionErreurPDT(iRes, "TID_bChoixSsanOuSaisieEAN128")
                End If

            End If



        End While

    End Function

    'Saisie de l'article
    Private Function TID_bSaisieArticle(ByRef vsArticle As String, ByRef vsArticleLibelle As String, ByRef vsArticleType As String, ByRef vbFinSaisie As Boolean) As Boolean

        Dim sScan As String = ""
        Dim iRes As Integer = 0
        TID_bSaisieArticle = False

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not TID_bSaisieArticle And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            TID_AffichageSaisieARTICLE()

            'Demande de saisie
            sScan = go_IO.RFInput("", 15, CHR_nCentrer(15), 4, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else

                        vsArticle = Trim(sScan)
                        If (API_TID_bRechercheArticle(vsArticle, vsArticleLibelle, vsArticleType)) Then
                            TID_bSaisieArticle = True
                        End If

                    End If
                Else
                    vbFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "TID_bSaisieArticle")
            End If

        End While

    End Function

    'Saisie du lot
    Private Function TID_bSaisieLot(ByVal vsEmplacementdeDebut As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByRef vsLot As String, ByRef vsNumOrdre As String, ByRef vbFinSaisie As Boolean) As Boolean

        Dim sScan As String = ""
        Dim iRes As Integer = 0
        TID_bSaisieLot = False
        vsLot = ""
        vsNumOrdre = ""

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not TID_bSaisieLot And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            TID_AffichageSaisieLOT(vsEmplacementdeDebut, vsArticle, vsArticleLibelle)

            'Demande de saisie
            sScan = go_IO.RFInput("", 20, CHR_nCentrer(20), 6, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else

                        vsLot = Trim(sScan)
                        If (API_TID_bControleArticleLot(vsNumOrdre, vsArticle, vsLot)) Then
                            TID_bSaisieLot = True
                        End If

                    End If
                Else
                    vbFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "TID_bSaisieLot")
            End If

        End While

    End Function

    'Saisie de la quantité de la palette
    Private Function TID_bSaisieQuantite(ByVal vsDepot As String, ByVal vsEmplacement As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByVal vsLot As String, ByVal vsNumOrdre As String,
                                         ByRef vsQuantite As String, ByRef vbFinSaisie As Boolean) As Boolean

        Dim sScan As String = ""
        Dim iRes As Integer = 0
        Dim nQuantite As Long = 0

        TID_bSaisieQuantite = False
        vsQuantite = ""

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not TID_bSaisieQuantite And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            TID_AffichageSaisieQUANTITE(vsEmplacement, vsArticle, vsArticleLibelle, vsLot, vsNumOrdre)

            'Demande de saisie
            sScan = go_IO.RFInput("", 10, CHR_nCentrer(10), 8, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
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
                                If (API_TID_bControleQuantiteArticleLot(vsDepot, vsArticle, vsLot, vsQuantite, vsEmplacement)) Then
                                    TID_bSaisieQuantite = True
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
                WRL_GestionErreurPDT(iRes, "TID_bSaisieQuantite")
            End If

        End While

    End Function

    'Saisie de la quantité préparée
    Private Function TID_bSaisieQuantitePreparee(ByVal vsIndexDeLivraison As String, ByVal vsNumLP As String, ByVal vsDepot As String, ByVal vsEmplacement As String, ByVal vsArticle As String,
                                                 ByVal vsArticleLibelle As String, ByVal vsLot As String, ByVal vsNumOrdre As String, ByRef vsQuantite As String, ByVal vsQuantiteRestanteAPreparer As String,
                                                 ByRef vbFinSaisie As Boolean) As Boolean

        Dim sScan As String = ""
        Dim iRes As Integer = 0
        Dim nQuantite As Long = 0
        Dim nQuantiteRestanteAPreparer As Long = 0

        TID_bSaisieQuantitePreparee = False

        If (IsNumeric(CHR_TransformeSeparateurPourNumerique(Trim(vsQuantiteRestanteAPreparer)))) Then
            nQuantiteRestanteAPreparer = (CHR_TransformeSeparateurPourNumerique(Trim(vsQuantiteRestanteAPreparer)))
        End If


        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not TID_bSaisieQuantitePreparee And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            TID_AffichageSaisieQUANTITEPreparee(vsIndexDeLivraison, vsNumLP, vsEmplacement, vsArticle, vsArticleLibelle, vsLot, vsQuantiteRestanteAPreparer)

            'Demande de saisie
            sScan = go_IO.RFInput(vsQuantite, 10, CHR_nCentrer(10), 9, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else
                        If (IsNumeric(CHR_TransformeSeparateurPourNumerique(Trim(sScan)))) Then
                            nQuantite = (CHR_TransformeSeparateurPourNumerique(Trim(sScan)))
                            If (nQuantite <> 0) Then
                                If (API_TID_bControleQuantiteArticleLot(vsDepot, vsArticle, vsLot, Trim(sScan), vsEmplacement)) Then

                                    If (nQuantite > nQuantiteRestanteAPreparer) Then
                                        MSG_AfficheErreur(giERR_QUANTITE_TROP_GRANDE)
                                    Else
                                        vsQuantite = Trim(sScan)
                                        TID_bSaisieQuantitePreparee = True
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
                WRL_GestionErreurPDT(iRes, "TID_bSaisieQuantitePreparee")
            End If

        End While

    End Function

    'Affichage du titre pour saisie de l'index de livraison et du N° de LP
    Private Sub TID_AffichageSaisieLP()
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 2, CHR_sCentrer("NUMERO LP"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage de l'écran choix de la saisie (Scan CAB ou saisie manuelle)
    Private Sub TID_AffichageChoix(ByVal vsIndexDeLivraison As String, ByVal vsNumLP As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByVal vsQuantite As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 1, "INDEX: " & vsIndexDeLivraison & " LP: " & vsNumLP, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 2, "ART. : " & vsArticle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, vsArticleLibelle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 4, "QTE : " & vsQuantite, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 6, CHR_sCentrer("CAB PALETTE"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage de l'écran de validation de la LP
    Private Sub TID_AffichageValidationLP(ByVal vsIndexDeLivraison As String, ByVal vsNumLP As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 1, "INDEX: " & vsIndexDeLivraison & " LP: " & vsNumLP, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 4, CHR_sCentrer("VOULEZ VOUS VALIDER"), WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 5, CHR_sCentrer("LA SORTIE DES STOCK"), WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 6, CHR_sCentrer("DE CETTE LP"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage du titre pour saisie de l'article
    Private Sub TID_AffichageSaisieARTICLE()
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 3, CHR_sCentrer("ARTICLE"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage du titre pour saisie du lot
    Private Sub TID_AffichageSaisieLOT(ByVal vsEmplacementDeDebut As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 1, "EMPL. : " & vsEmplacementDeDebut, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 2, "ART. : " & vsArticle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, vsArticleLibelle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 5, CHR_sCentrer("LOT"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage du titre pour saisie de la quantité de la palette
    Private Sub TID_AffichageSaisieQUANTITE(ByVal vsEmplacement As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByVal vsLot As String, ByVal vsNumOrdre As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 1, "EMPL. : " & vsEmplacement, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 2, "ART. : " & vsArticle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, vsArticleLibelle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 4, "LOT : " & vsLot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 5, "N° ORDRE : " & vsNumOrdre, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 7, CHR_sCentrer("QUANTITE"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage du titre pour saisie de la quantité préparée
    Private Sub TID_AffichageSaisieQUANTITEPreparee(ByVal vsIndexDeLivraison As String, ByVal vsNumLP As String, ByVal vsEmplacement As String, ByVal vsArticle As String,
                                                    ByVal vsArticleLibelle As String, ByVal vsLot As String, ByVal vsQuantiteRestanteAPreparer As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 1, "INDEX: " & vsIndexDeLivraison & " / " & " LP: " & vsNumLP, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 2, "EMPL. : " & vsEmplacement, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, "ART. : " & vsArticle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 4, vsArticleLibelle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 5, "LOT : " & vsLot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 6, "QTE A PREPARER : " & vsQuantiteRestanteAPreparer, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 8, CHR_sCentrer("QUANTITE"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

End Module