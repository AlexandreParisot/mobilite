Option Strict Off
Option Explicit On
Module modEcranEEP
    Dim msTitre As String


    'Entrée de l'option EEP - Edition d'une étiquette palette soit en lisant un EAN128 soit en saisissant toutes les informations utiles à la constitution du code à barre correspondant.
    ' Le code à barre contient : N° d'ordre (AO ou OF) non obligatoire, code article MMS001, N° de lot, Quantité sur la palette
    Public Sub EcranEEP(ByRef vsTitre As String)
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

        msTitre = vsTitre

        If EEP_bSaisieDepot(sDepot, sDepotLibelle) Then

            If EEP_bSaisieEmplacement(sDepot, sEmplacement) Then

                While Not bFinSaisie

                    sEAN128Saisi = ""
                    sArticle = ""
                    sLot = ""
                    sQuantite = ""
                    EEP_bChoixScanOuSaisieEAN128(sDepot, sEmplacement, sEAN128Saisi, bFinSaisie)

                    If (Not bFinSaisie) Then

                        ' ====================
                        'Le CAB a été scanné ou saisi
                        If (sEAN128Saisi <> "") Then
                            If (CHR_bRecupInfoEAN128(sEAN128Saisi, sNumOrdre, sArticle, sLot, sQuantite)) Then

                                If (API_EEP_bControleArticleLot(sNumOrdre, sArticle, sLot)) Then

                                    If (API_EEP_bControleQuantiteArticleLot(sDepot, sArticle, sLot, sQuantite, sEmplacement)) Then

                                        API_EEP_bEditionEtiquettePalette(sDepot, sEmplacement, sNumOrdre, sArticle, sLot, sQuantite)

                                    End If

                                End If

                            End If

                        Else
                            ' ====================
                            'Les éléments du CAB doivent être saisis Article + Lot + Quantité
                            If (EEP_bSaisieArticle(sDepot, sEmplacement, sArticle, sArticleLibelle, bFinSaisie)) Then

                                If (EEP_bSaisieLot(sDepot, sEmplacement, sArticle, sArticleLibelle, sLot, sNumOrdre, bFinSaisie)) Then

                                    If (EEP_bSaisieQuantite(sDepot, sEmplacement, sArticle, sArticleLibelle, sLot, sNumOrdre, sQuantite, bFinSaisie)) Then

                                        API_EEP_bEditionEtiquettePalette(sDepot, sEmplacement, sNumOrdre, sArticle, sLot, sQuantite)

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
    Private Function EEP_bSaisieDepot(ByRef vsDepot As String, ByRef vsDepotLibelle As String) As Boolean

        EEP_bSaisieDepot = False
        Dim bFinSaisie As Boolean = False
        Dim sScan As String = ""
        Dim iRes As Integer = 0

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not EEP_bSaisieDepot And Not gbErreurCommunication And Not bFinSaisie

            'Affichage
            EEP_AffichageSaisieDEPOT(vsDepot)

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
                            EEP_bSaisieDepot = True
                        End If

                    End If
                Else
                    bFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "EEP_bSaisieDepot")
            End If

        End While

    End Function

    'Saisie de l'emplacement
    Private Function EEP_bSaisieEmplacement(ByVal vsDepot As String, ByRef vsEmplacement As String) As Boolean

        EEP_bSaisieEmplacement = False
        Dim bFinSaisie As Boolean = False
        Dim sScan As String = ""
        Dim iRes As Integer = 0

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not EEP_bSaisieEmplacement And Not gbErreurCommunication And Not bFinSaisie

            'Affichage
            EEP_AffichageSaisieEMPLACEMENT(vsDepot, vsEmplacement)

            'Demande de saisie
            sScan = go_IO.RFInput("", 10, CHR_nCentrer(10), 4, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        bFinSaisie = True
                    Else

                        vsEmplacement = Trim(sScan)
                        If (API_EEP_bRechercheEmplacement(vsDepot, vsEmplacement)) Then
                            EEP_bSaisieEmplacement = True
                        End If

                    End If
                Else
                    bFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "EEP_bSaisieEmplacement")
            End If

        End While

    End Function

    'Choix, soit Scan ou Saisie du code EAN128
    Private Function EEP_bChoixScanOuSaisieEAN128(ByVal vsDepot As String, ByVal vsEmplacement As String, ByRef vsEAN128Saisi As String, ByRef vbFinSaisie As Boolean) As Boolean

        EEP_bChoixScanOuSaisieEAN128 = False
        Dim sScan As String = ""
        Dim iRes As Integer = 0
        vsEAN128Saisi = ""

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not EEP_bChoixScanOuSaisieEAN128 And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            EEP_AffichageChoix(vsDepot, vsEmplacement)

            'Demande de saisie
            sScan = go_IO.RFInput("", 70, CHR_nCentrer(70), 6, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT_SAISIE, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else
                        If sScan = gCST_TOUCHE_SAISIE_F5 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                            EEP_bChoixScanOuSaisieEAN128 = True
                        Else
                            vsEAN128Saisi = sScan
                            EEP_bChoixScanOuSaisieEAN128 = True
                        End If
                    End If
                Else
                    vbFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "EEP_bChoixSsanOuSaisieEAN128")
            End If
        End While

    End Function

    'Saisie de l'article
    Private Function EEP_bSaisieArticle(ByVal vsDepot As String, ByVal vsEmplacement As String, ByRef vsArticle As String, ByRef vsArticleLibelle As String, ByRef vbFinSaisie As Boolean) As Boolean

        EEP_bSaisieArticle = False
        Dim sScan As String = ""
        Dim iRes As Integer = 0

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not EEP_bSaisieArticle And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            EEP_AffichageSaisieARTICLE(vsDepot, vsEmplacement)

            'Demande de saisie
            sScan = go_IO.RFInput("", 15, CHR_nCentrer(15), 6, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else

                        vsArticle = Trim(sScan)
                        If (API_EEP_bRechercheArticle(vsArticle, vsArticleLibelle)) Then
                            EEP_bSaisieArticle = True
                        End If

                    End If
                Else
                    vbFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "EEP_bSaisieArticle")
            End If

        End While

    End Function

    'Saisie du lot
    Private Function EEP_bSaisieLot(ByVal vsDepot As String, ByVal vsEmplacement As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByRef vsLot As String, ByRef vsNumOrdre As String, ByRef vbFinSaisie As Boolean) As Boolean

        EEP_bSaisieLot = False
        Dim sScan As String = ""
        Dim iRes As Integer = 0
        vsLot = ""
        vsNumOrdre = ""

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not EEP_bSaisieLot And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            EEP_AffichageSaisieLOT(vsDepot, vsEmplacement, vsArticle, vsArticleLibelle)

            'Demande de saisie
            sScan = go_IO.RFInput("", 20, CHR_nCentrer(20), 8, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else

                        vsLot = Trim(sScan)
                        If (API_EEP_bControleArticleLot(vsNumOrdre, vsArticle, vsLot)) Then
                            EEP_bSaisieLot = True
                        End If

                    End If
                Else
                    vbFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "EEP_bSaisieLot")
            End If

        End While

    End Function

    'Saisie de la quantité
    Private Function EEP_bSaisieQuantite(ByVal vsDepot As String, ByVal vsEmplacement As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByVal vsLot As String, ByVal vsNumOrdre As String,
                                         ByRef vsQuantite As String, ByRef vbFinSaisie As Boolean) As Boolean

        EEP_bSaisieQuantite = False
        Dim sScan As String = ""
        Dim iRes As Integer = 0
        Dim nQuantite As Long = 0
        vsQuantite = ""

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not EEP_bSaisieQuantite And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            EEP_AffichageSaisieQUANTITE(vsDepot, vsEmplacement, vsArticle, vsArticleLibelle, vsLot, vsNumOrdre)

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
                                If (API_EEP_bControleQuantiteArticleLot(vsDepot, vsArticle, vsLot, vsQuantite, vsEmplacement)) Then
                                    EEP_bSaisieQuantite = True
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
                WRL_GestionErreurPDT(iRes, "EEP_bSaisieQuantite")
            End If

        End While

    End Function

    'Affichage du titre pour saisie du dépôt
    Private Sub EEP_AffichageSaisieDEPOT(ByVal vsDepot As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 2, CHR_sCentrer("DEPOT"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage du titre pour saisie de l'emplacement
    Private Sub EEP_AffichageSaisieEMPLACEMENT(ByVal vsDepot As String, ByVal vsEmplacement As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 2, "DEPOT : " & vsDepot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, CHR_sCentrer("EMPLACEMENT"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage de l'écran choix de la saisie (Scan CAB ou saisie manuelle)
    Private Sub EEP_AffichageChoix(ByVal vsDepot As String, ByVal vsEmplacement As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 2, "DEPOT : " & vsDepot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, "EMPL. : " & vsEmplacement, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 5, CHR_sCentrer("CAB PALETTE"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage du titre pour saisie de l'article
    Private Sub EEP_AffichageSaisieARTICLE(ByVal vsDepot As String, ByVal vsEmplacement As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 2, "DEPOT : " & vsDepot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, "EMPL. : " & vsEmplacement, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 5, CHR_sCentrer("ARTICLE"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage du titre pour saisie du lot
    Private Sub EEP_AffichageSaisieLOT(ByVal vsDepot As String, ByVal vsEmplacement As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 2, "DEPOT : " & vsDepot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, "EMPL. : " & vsEmplacement, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 4, "ART. : " & vsArticle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 5, vsArticleLibelle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 7, CHR_sCentrer("LOT"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage du titre pour saisie de la quantité
    Private Sub EEP_AffichageSaisieQUANTITE(ByVal vsDepot As String, ByVal vsEmplacement As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByVal vsLot As String, ByVal vsNumOrdre As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 2, "DEPOT : " & vsDepot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, "EMPL. : " & vsEmplacement, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 4, "ART. : " & vsArticle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 5, vsArticleLibelle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 6, "LOT : " & vsLot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 7, "N° ORDRE : " & vsNumOrdre, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 9, CHR_sCentrer("QUANTITE"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

End Module