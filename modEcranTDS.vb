Option Strict Off
Option Explicit On
Module modEcranTDS
    Dim msTitre As String

    'Entrée de l'option TDS - Transfert de stock
    Public Sub EcranTDS(ByRef vsTitre As String)
        Dim bFinSaisie As Boolean = False
        Dim sEmplacementDeDebut As String = ""
        Dim sEAN128Saisi As String = ""
        Dim sEmplacementDeFin As String = ""
        Dim sArticle As String = ""
        Dim sArticleLibelle As String = ""
        Dim sLot As String = ""
        Dim sQuantite As String = ""
        Dim sDateCourante As String = ""
        Dim sHeureCourante As String = ""
        Dim sTypeEmplacementDeDebut As String = ""
        Dim sStatutIDStock As String = ""
        Dim sReferenceLot2 As String = ""
        Dim sEmplacementArticleDepot As String = ""
        Dim sNumOrdre As String = ""
        Dim sQuantiteAffectee As String = ""
        Dim sDatePeremption As String = ""
        Dim bSuiteTraitementOK As Boolean = False


        msTitre = vsTitre

        While Not bFinSaisie

            TDS_bSaisieEmplacementDeDebut(sEmplacementDeDebut, sTypeEmplacementDeDebut, bFinSaisie)

            If (Not bFinSaisie) Then

                sEAN128Saisi = ""
                sArticle = ""
                sLot = ""
                sQuantite = ""
                TDS_bChoixScanOuSaisieEAN128(sEmplacementDeDebut, sEAN128Saisi, bFinSaisie)

                If (Not bFinSaisie) Then

                    ' ====================
                    'Le CAB a été scanné ou saisi
                    If (sEAN128Saisi <> "") Then
                        If (CHR_bRecupInfoEAN128(sEAN128Saisi, sNumOrdre, sArticle, sLot, sQuantite)) Then

                            If (API_TDS_bControleArticleLot(sNumOrdre, sArticle, sLot)) Then

                                If (API_TDS_bControleQuantiteArticleLot(sArticle, sLot, sQuantite, sEmplacementDeDebut, sStatutIDStock, sQuantiteAffectee, sDatePeremption)) Then

                                    bSuiteTraitementOK = True

                                End If

                            End If

                        End If

                    Else
                        ' ====================
                        'Les éléments du CAB doivent être saisis Article + Lot
                        sArticle = ""
                        sLot = ""
                        sQuantite = ""
                        If (TDS_bSaisieArticle(sEmplacementDeDebut, sArticle, sArticleLibelle, bFinSaisie)) Then

                            If (TDS_bSaisieLot(sEmplacementDeDebut, sArticle, sArticleLibelle, sLot, sNumOrdre, bFinSaisie)) Then

                                If (TDS_bSaisieQuantite(sEmplacementDeDebut, sArticle, sArticleLibelle, sLot, sNumOrdre, sQuantite, bFinSaisie)) Then

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

                If (TDS_bSaisieEmplacementDeFin(sEmplacementDeDebut, sArticle, sArticleLibelle, sLot, sQuantite, sEmplacementArticleDepot, sEmplacementDeFin, bFinSaisie)) Then

                    API_bRecupDateHeure(sDateCourante, sHeureCourante)

                    If (API_TDS_bTransfertIdStock(sEmplacementDeDebut, sArticle, sLot, sQuantite, sEmplacementDeFin)) Then
                    End If

                End If

            End If

        End While

    End Sub

    'Saisie de l'emplacement de début
    Private Function TDS_bSaisieEmplacementDeDebut(ByRef vsEmplacementDeDebut As String, ByRef vsTypeEmplacementDeDebut As String, ByRef vbFinSaisie As Boolean) As Boolean

        Dim sScan As String = ""
        Dim iRes As Integer = 0

        TDS_bSaisieEmplacementDeDebut = False
        vbFinSaisie = False

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not TDS_bSaisieEmplacementDeDebut And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            TDS_AffichageSaisieEMPLACEMENTdeDEBUT()

            'Demande de saisie
            sScan = go_IO.RFInput("", 10, CHR_nCentrer(10), 3, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreu
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else

                        vsEmplacementDeDebut = Trim(sScan)
                        If (API_TDS_bRechercheEmplacement(vsEmplacementDeDebut, vsTypeEmplacementDeDebut)) Then
                            TDS_bSaisieEmplacementDeDebut = True
                        End If

                    End If
                Else
                    vbFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "TDS_bSaisieEmplacementDeDebut")
            End If

        End While

    End Function

    'Choix, soit Scan ou Saisie du code EAN128
    Private Function TDS_bChoixScanOuSaisieEAN128(ByVal vsEmplacement As String, ByRef vsEAN128Saisi As String, ByRef vbFinSaisie As Boolean) As Boolean

        Dim sScan As String = ""
        Dim iRes As Integer = 0

        TDS_bChoixScanOuSaisieEAN128 = False
        vsEAN128Saisi = ""

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not TDS_bChoixScanOuSaisieEAN128 And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            TDS_AffichageChoix(vsEmplacement)

            'Demande de saisie
            sScan = go_IO.RFInput("", 70, CHR_nCentrer(70), 4, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT_SAISIE, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else
                        If sScan = gCST_TOUCHE_SAISIE_F5 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                            TDS_bChoixScanOuSaisieEAN128 = True
                        Else
                            vsEAN128Saisi = sScan
                            TDS_bChoixScanOuSaisieEAN128 = True
                        End If
                    End If
                Else
                    vbFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "TDS_bChoixSsanOuSaisieEAN128")
            End If
        End While

    End Function

    'Saisie de l'article
    Private Function TDS_bSaisieArticle(ByVal vsEmplacementDeDebut As String, ByRef vsArticle As String, ByRef vsArticleLibelle As String, ByRef vbFinSaisie As Boolean) As Boolean

        Dim sScan As String = ""
        Dim iRes As Integer = 0
        TDS_bSaisieArticle = False

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not TDS_bSaisieArticle And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            TDS_AffichageSaisieARTICLE(vsEmplacementDeDebut)

            'Demande de saisie
            sScan = go_IO.RFInput("", 15, CHR_nCentrer(15), 4, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else

                        vsArticle = Trim(sScan)
                        If (API_TDS_bRechercheArticle(vsArticle, vsArticleLibelle)) Then
                            TDS_bSaisieArticle = True
                        End If

                    End If
                Else
                    vbFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "TDS_bSaisieArticle")
            End If

        End While

    End Function

    'Saisie du lot
    Private Function TDS_bSaisieLot(ByVal vsEmplacementdeDebut As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByRef vsLot As String, ByRef vsNumOrdre As String, ByRef vbFinSaisie As Boolean) As Boolean

        Dim sScan As String = ""
        Dim iRes As Integer = 0
        TDS_bSaisieLot = False
        vsLot = ""
        vsNumOrdre = ""

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not TDS_bSaisieLot And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            TDS_AffichageSaisieLOT(vsEmplacementdeDebut, vsArticle, vsArticleLibelle)

            'Demande de saisie
            sScan = go_IO.RFInput("", 20, CHR_nCentrer(20), 6, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else

                        vsLot = Trim(sScan)
                        If (API_TDS_bControleArticleLot(vsNumOrdre, vsArticle, vsLot)) Then
                            TDS_bSaisieLot = True
                        End If

                    End If
                Else
                    vbFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "TDS_bSaisieLot")
            End If

        End While

    End Function

    'Saisie de la quantité
    Private Function TDS_bSaisieQuantite(ByVal vsEmplacement As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByVal vsLot As String, ByVal vsNumOrdre As String,
                                         ByRef vsQuantite As String, ByRef vbFinSaisie As Boolean) As Boolean

        Dim sScan As String = ""
        Dim iRes As Integer = 0
        Dim nQuantite As Long = 0

        TDS_bSaisieQuantite = False
        vsQuantite = ""

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not TDS_bSaisieQuantite And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            TDS_AffichageSaisieQUANTITE(vsEmplacement, vsArticle, vsArticleLibelle, vsLot, vsNumOrdre)

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
                                If (API_TDS_bControleQuantiteArticleLot(vsArticle, vsLot, vsQuantite, vsEmplacement)) Then
                                    TDS_bSaisieQuantite = True
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
                WRL_GestionErreurPDT(iRes, "TDS_bSaisieQuantite")
            End If

        End While

    End Function

    'Saisie de l'emplacement de fin
    Private Function TDS_bSaisieEmplacementDeFin(ByVal vsEmplacementDeDebut As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByVal vsLot As String, ByVal vsQuantite As String,
                                                 ByVal vsEmplacementArticleDepot As String, ByRef vsEmplacementDeFin As String, ByRef vbFinSaisie As Boolean) As Boolean

        Dim sScan As String = ""
        Dim iRes As Integer = 0

        TDS_bSaisieEmplacementDeFin = False
        vsEmplacementDeFin = ""

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not TDS_bSaisieEmplacementDeFin And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            TDS_AffichageSaisieEMPLACEMENTdeFIN(vsEmplacementDeDebut, vsArticle, vsArticleLibelle, vsLot, vsQuantite)

            'Demande de saisie
            sScan = go_IO.RFInput(vsEmplacementArticleDepot, 10, CHR_nCentrer(10), 8, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else

                        vsEmplacementDeFin = Trim(sScan)
                        If (API_TDS_bRechercheEmplacement(vsEmplacementDeFin)) Then
                            TDS_bSaisieEmplacementDeFin = True
                        End If

                    End If
                Else
                    vbFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "TDS_bSaisieEmplacementDeFin")
            End If

        End While

    End Function

    'Affichage du titre pour saisie de l'Emplacement de début
    Private Sub TDS_AffichageSaisieEMPLACEMENTdeDEBUT()
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 2, CHR_sCentrer("EMPLACEMENT DEBUT"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage de l'écran choix de la saisie (Scan CAB ou saisie manuelle)
    Private Sub TDS_AffichageChoix(ByVal vsEmplacementDeDebut As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 1, "EMPL. : " & vsEmplacementDeDebut, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, CHR_sCentrer("CAB PALETTE"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage du titre pour saisie de l'article
    Private Sub TDS_AffichageSaisieARTICLE(ByVal vsEmplacementDeDebut As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 1, "EMPL. : " & vsEmplacementDeDebut, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, CHR_sCentrer("ARTICLE"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage du titre pour saisie du lot
    Private Sub TDS_AffichageSaisieLOT(ByVal vsEmplacementDeDebut As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 1, "EMPL. : " & vsEmplacementDeDebut, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 2, "ART. : " & vsArticle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, vsArticleLibelle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 5, CHR_sCentrer("LOT"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage du titre pour saisie de la quantité
    Private Sub TDS_AffichageSaisieQUANTITE(ByVal vsEmplacement As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByVal vsLot As String, ByVal vsNumOrdre As String)
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

    'Affichage du titre pour saisie de l'Emplacement de fin
    Private Sub TDS_AffichageSaisieEMPLACEMENTdeFIN(ByVal vsEmplacementDeDebut As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByVal vsLot As String, ByVal vsQuantite As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 1, "EMPL. : " & vsEmplacementDeDebut, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 2, "ART. : " & vsArticle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, vsArticleLibelle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 4, "LOT : " & vsLot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 5, "QTE : " & vsQuantite, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 7, CHR_sCentrer("EMPLACEMENT FIN"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

End Module