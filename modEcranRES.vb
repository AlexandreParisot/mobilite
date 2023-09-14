Option Strict Off
Option Explicit On
Module modEcranRES
    Dim msTitre As String

    'Entrée de l'option RES - Retour en stock depuis le bord de chaîne
    Public Sub EcranRES(ByRef vsTitre As String)
        Dim bFinSaisie As Boolean = False
        Dim sEmplacementDeDebut As String = ""
        Dim sEmplacementDeFin As String = ""
        Dim sEmplacementFinalParDefaut As String = ""
        Dim sArticle As String = ""
        Dim sArticleLibelle As String = ""
        Dim sArticleType As String = ""
        Dim sLot As String = ""
        Dim sQuantite As String = ""
        Dim sDateCourante As String = ""
        Dim sHeureCourante As String = ""
        Dim sTypeEmplacementDeDebut As String = ""
        Dim sNouveauLot As String = ""
        Dim sStatutIDStock As String = ""
        Dim sReferenceLot2 As String = ""
        Dim sEmplacementArticleDepot As String = ""


        msTitre = vsTitre


        While Not bFinSaisie

            RES_bSaisieEmplacementDeDebut(sEmplacementDeDebut, sTypeEmplacementDeDebut, bFinSaisie)

            If (Not bFinSaisie) Then

                If (RES_bRechercheDernierArticleSurEmplacementDeDebut(sEmplacementDeDebut, sArticle, sArticleLibelle, sArticleType, bFinSaisie)) Then

                    If (RES_bRechercheDernierLotSurEmplacementDeDebut(sTypeEmplacementDeDebut, sEmplacementDeDebut, sArticle, sArticleLibelle, sLot, sReferenceLot2, sQuantite, sStatutIDStock, bFinSaisie)) Then

                        If (RES_bSaisieQuantiteIDStockSurEmplacementDeDebut(sEmplacementDeDebut, sArticle, sArticleLibelle, sLot, sReferenceLot2, sQuantite, sStatutIDStock, bFinSaisie)) Then

                            API_RES_bRechercheArticleDepot(sArticle, sEmplacementArticleDepot)

                            If (sArticleType.StartsWith("TB")) Then
                                sEmplacementFinalParDefaut = gTab_Configuration.sRES_WHSL_TB_Final
                            Else
                                sEmplacementFinalParDefaut = sEmplacementArticleDepot
                            End If

                            If (RES_bSaisieEmplacementDeFin(sEmplacementDeDebut, sArticle, sArticleLibelle, sLot, sQuantite, sEmplacementFinalParDefaut, sReferenceLot2, sEmplacementDeFin, bFinSaisie)) Then

                                API_bRecupDateHeure(sDateCourante, sHeureCourante)

                                If (sTypeEmplacementDeDebut = gTab_Configuration.sSLPT_Bord_De_Chaine_Normal) Then

                                    If (sReferenceLot2 <> "") Then
                                        sNouveauLot = sReferenceLot2
                                        API_RES_bReclassIdStock(sEmplacementDeDebut, sArticle, sLot, sQuantite, sNouveauLot, sStatutIDStock)
                                        sLot = sNouveauLot
                                    End If

                                ElseIf (sTypeEmplacementDeDebut = gTab_Configuration.sSLPT_Bord_De_Chaine_Bib) Then

                                    If (sReferenceLot2 <> "") Then
                                        sNouveauLot = sReferenceLot2
                                        API_RES_bReclassIdStock(sEmplacementDeDebut, sArticle, sLot, sQuantite, sNouveauLot, sStatutIDStock)
                                        sLot = sNouveauLot
                                    End If

                                End If

                                If (API_RES_bTransfertIdStock(sEmplacementDeDebut, sArticle, sLot, sQuantite, sEmplacementDeFin)) Then
                                    API_RES_bEditionEtiquettePalette(sEmplacementDeFin, "", sArticle, sLot, sQuantite)
                                End If

                            End If

                        End If

                    End If

                End If

            End If

        End While

    End Sub

    'Saisie de l'emplacement de début
    Private Function RES_bSaisieEmplacementDeDebut(ByRef vsEmplacementDeDebut As String, ByRef vsTypeEmplacementDeDebut As String, ByRef vbFinSaisie As Boolean) As Boolean

        Dim sScan As String = ""
        Dim iRes As Integer = 0

        RES_bSaisieEmplacementDeDebut = False
        vbFinSaisie = False

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not RES_bSaisieEmplacementDeDebut And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            RES_AffichageSaisieEMPLACEMENTdeDEBUT()

            'Demande de saisie
            sScan = go_IO.RFInput("", 10, CHR_nCentrer(10), 3, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT_IMPRIME_DERNIERE_ETIQUETTE, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else

                        If sScan = gCST_TOUCHE_IMPRIME_DERNIERE_ETIQUETTE_F2 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                            API_RES_bImpression_Derniere_Etiquette()
                        Else

                            vsEmplacementDeDebut = Trim(sScan)
                            If (API_RES_bRechercheEmplacement(vsEmplacementDeDebut, vsTypeEmplacementDeDebut)) Then
                                If (vsTypeEmplacementDeDebut <> gTab_Configuration.sSLPT_Bord_De_Chaine_Normal And vsTypeEmplacementDeDebut <> gTab_Configuration.sSLPT_Bord_De_Chaine_Bib) Then
                                    MSG_AfficheErreur(giERR_TYPE_EMPLACEMENT_INVALIDE)
                                Else
                                    RES_bSaisieEmplacementDeDebut = True
                                End If

                            End If

                        End If

                    End If
                Else
                    vbFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "RES_bSaisieEmplacementDeDebut")
            End If

        End While

    End Function

    'Recherche des derniers articles déposés sur l'emplacement de début 
    Private Function RES_bRechercheDernierArticleSurEmplacementDeDebut(ByVal vsEmplacementDeDebut As String, ByRef vsArticle As String, ByRef vsArticleLibelle As String, ByRef vsArticleType As String,
                                                                       ByRef vbFinSaisie As Boolean) As Boolean

        Dim sScan As String = ""
        Dim iRes As Integer = 0
        Dim sTablo_Article(0)
        Dim sTablo_Lot(0)
        Dim sTablo_BRE2(0)
        Dim sTablo_Qte(0)
        Dim nIndex As Short = 0


        RES_bRechercheDernierArticleSurEmplacementDeDebut = False
        vbFinSaisie = False
        vsArticle = ""
        vsArticleLibelle = ""
        vsArticleType = ""


        API_RES_bRechercheDernierIdStockSurEmplacement(vsEmplacementDeDebut, vsArticle, "", sTablo_Article, sTablo_Lot, sTablo_BRE2, sTablo_Qte)

        'Si on a pas trouvé d'article sur l'emplacement on quitte le scénario
        If (sTablo_Article.Length > 0) Then
            If (sTablo_Article(nIndex) = "") Then
                vbFinSaisie = True
            End If
        Else
            vbFinSaisie = True
        End If


        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not RES_bRechercheDernierArticleSurEmplacementDeDebut And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            RES_AffichageArticleIDStock(vsEmplacementDeDebut, sTablo_Article)

            'Demande de saisie
            sScan = go_IO.RFInput("", 1, CHR_nCentrer(1), 9, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else
                        If (IsNumeric(CHR_TransformeSeparateurPourNumerique(Trim(sScan)))) Then
                            nIndex = (CHR_TransformeSeparateurPourNumerique(Trim(sScan)))

                            vsArticle = sTablo_Article(nIndex - 1)
                            API_RES_bRechercheArticle(vsArticle, vsArticleLibelle, vsArticleType)
                            RES_bRechercheDernierArticleSurEmplacementDeDebut = True

                        Else
                            MSG_AfficheErreur(giERR_FORMAT_NUMERIC)
                        End If
                    End If
                Else
                    vbFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "RES_bSaisieArticle")
            End If

        End While

    End Function

    'Recherche des derniers lot de article déposés sur l'emplacement de début 
    Private Function RES_bRechercheDernierLotSurEmplacementDeDebut(ByVal vsTypeEmplacementDeDebut As String, ByVal vsEmplacementDeDebut As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByRef vsLot As String,
                                                                   ByRef vsReferenceLot2 As String, ByRef vsQuantite As String, ByRef vsStatutIDStock As String, ByRef vbFinSaisie As Boolean) As Boolean

        Dim sScan As String = ""
        Dim iRes As Integer = 0
        Dim sTablo_Article(0)
        Dim sTablo_Lot(0)
        Dim sTablo_BRE2(0)
        Dim sTablo_Qte(0)
        Dim nIndex As Short = 0
        Dim sEAN128Saisi As String = ""
        Dim sEAN128_NumOrdre As String = ""
        Dim sEAN128_Article As String = ""
        Dim sEAN128_Lot As String = ""
        Dim sEAN128_Quantite As String = ""


        RES_bRechercheDernierLotSurEmplacementDeDebut = False
        vbFinSaisie = False
        vsLot = ""
        vsReferenceLot2 = ""
        vsQuantite = ""
        vsStatutIDStock = ""


        ' Si on est sur un emplacement bord de chaîne "normal", on recherche les derniers lots déposés sur cet emplacement

        API_RES_bRechercheDernierIdStockSurEmplacement(vsEmplacementDeDebut, vsArticle, vsLot, sTablo_Article, sTablo_Lot, sTablo_BRE2, sTablo_Qte)

        'Si on a pas trouvé d'article sur l'emplacement on quitte le scénario
        If (sTablo_Article.Length > 0) Then
            If (sTablo_Article(nIndex) = "") Then
                vbFinSaisie = True
            End If
        Else
            vbFinSaisie = True
        End If

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not RES_bRechercheDernierLotSurEmplacementDeDebut And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            RES_AffichageLotIDStock(vsTypeEmplacementDeDebut, vsEmplacementDeDebut, vsArticle, vsArticleLibelle, sTablo_Lot(nIndex), sTablo_BRE2(nIndex))

            'Demande de saisie
            sScan = go_IO.RFInput(sTablo_Lot(nIndex), 20, CHR_nCentrer(20), 9, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT_SUIVANT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else

                        If sScan = gCST_TOUCHE_SUIVANT_F2 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then

                            If (sTablo_Article.Length - 1 >= nIndex + 1) Then
                                nIndex = nIndex + 1
                                If (sTablo_Article(nIndex) = "") Then
                                    'On averti l'utilisateur, plus aucun lot à afficher 
                                    MSG_AfficheErreur(giERR_ID_STOCK_NON_TROUVE)
                                End If
                            Else
                                'On averti l'utilisateur, plus aucun lot à afficher 
                                MSG_AfficheErreur(giERR_ID_STOCK_NON_TROUVE)
                            End If

                        Else
                            vsLot = sTablo_Lot(nIndex)
                            vsReferenceLot2 = sTablo_BRE2(nIndex)
                            vsQuantite = sTablo_Qte(nIndex)
                            RES_bRechercheDernierLotSurEmplacementDeDebut = True
                        End If

                    End If
                Else
                    vbFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "RES_bRechercheDernierLotSurEmplacementDeDebut_1")
            End If

        End While



        '---------- CAS DES BIB ----------
        If (RES_bRechercheDernierLotSurEmplacementDeDebut = True And vsTypeEmplacementDeDebut = gTab_Configuration.sSLPT_Bord_De_Chaine_Bib) Then
            ' Si on est sur un emplacement bord de chaîne "BIB", on demande à l'utilisateur de saisir le lot à retourner en stock

            RES_bRechercheDernierLotSurEmplacementDeDebut = False

            While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not RES_bRechercheDernierLotSurEmplacementDeDebut And Not gbErreurCommunication And Not vbFinSaisie

                'Affichage
                RES_AffichageLotASaisir(vsEmplacementDeDebut, vsArticle, vsArticleLibelle)

                'Demande de saisie
                sScan = go_IO.RFInput("", 70, CHR_nCentrer(70), 6, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT_SAISIE, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
                iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
                If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                    If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                        If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                            vbFinSaisie = True
                        Else

                            If sScan = gCST_TOUCHE_SAISIE_F5 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then

                                'Affichage
                                RES_AffichageLotASaisir(vsEmplacementDeDebut, vsArticle, vsArticleLibelle)
                                'Demande de saisie
                                sScan = go_IO.RFInput("", 20, CHR_nCentrer(20), 6, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
                                iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
                                If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                                    If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                                        If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                                            vbFinSaisie = True
                                        Else

                                            If (API_RES_bControleArticleLot(vsArticle, sScan)) Then
                                                vsReferenceLot2 = sScan
                                                RES_bRechercheDernierLotSurEmplacementDeDebut = True
                                            End If

                                        End If
                                    Else
                                        vbFinSaisie = True
                                    End If
                                Else
                                    WRL_GestionErreurPDT(iRes, "RES_bRechercheDernierLotSurEmplacementDeDebut_2")
                                End If

                            Else
                                sEAN128Saisi = sScan
                                If (CHR_bRecupInfoEAN128(sEAN128Saisi, sEAN128_NumOrdre, sEAN128_Article, sEAN128_Lot, sEAN128_Quantite)) Then

                                    If (sEAN128_Article <> vsArticle) Then
                                        MSG_AfficheErreur(giERR_ARTICLE_DIFFERENT)
                                    Else

                                        If (API_RES_bControleArticleLot(vsArticle, sEAN128_Lot)) Then
                                            vsReferenceLot2 = sEAN128_Lot
                                            RES_bRechercheDernierLotSurEmplacementDeDebut = True
                                        Else
                                            MSG_AfficheErreur(giERR_ID_STOCK_NON_TROUVE)
                                        End If

                                    End If

                                End If

                            End If

                        End If
                    Else
                        vbFinSaisie = True
                    End If
                Else
                    WRL_GestionErreurPDT(iRes, "RES_bRechercheDernierLotSurEmplacementDeDebut_3")
                End If

            End While

        End If

    End Function

    'Saisie de la quantité à retourner en stock 
    Private Function RES_bSaisieQuantiteIDStockSurEmplacementDeDebut(ByVal vsEmplacementDeDebut As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByVal vsLot As String,
                                                                     ByVal vsReferenceLot2 As String, ByRef vsQuantite As String, ByRef vsStatutIDStock As String, ByRef vbFinSaisie As Boolean) As Boolean

        Dim sScan As String = ""
        Dim iRes As Integer = 0
        Dim nQuantiteSaisie As Long = 0

        RES_bSaisieQuantiteIDStockSurEmplacementDeDebut = False
        vbFinSaisie = False
        vsStatutIDStock = ""

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not RES_bSaisieQuantiteIDStockSurEmplacementDeDebut And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            RES_AffichageIDStock(vsEmplacementDeDebut, vsArticle, vsArticleLibelle, vsLot, vsReferenceLot2, vsQuantite)

            'Demande de saisie
            sScan = go_IO.RFInput("", 10, CHR_nCentrer(10), 9, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else
                        If (IsNumeric(CHR_TransformeSeparateurPourNumerique(Trim(sScan)))) Then
                            nQuantiteSaisie = (CHR_TransformeSeparateurPourNumerique(Trim(sScan)))
                            If (nQuantiteSaisie <> 0) Then
                                If (API_RES_bControleQuantiteArticleLot(vsArticle, vsLot, Trim(sScan), vsEmplacementDeDebut, vsStatutIDStock, vsReferenceLot2)) Then
                                    vsQuantite = Trim(sScan)
                                    RES_bSaisieQuantiteIDStockSurEmplacementDeDebut = True
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
                WRL_GestionErreurPDT(iRes, "RES_bSaisieQuantité")
            End If

        End While

    End Function

    'Saisie de l'emplacement de fin
    Private Function RES_bSaisieEmplacementDeFin(ByVal vsEmplacementDeDebut As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByVal vsLot As String, ByVal vsQuantite As String,
                                                 ByVal vsEmplacementFinalParDefaut As String, ByVal vsReferenceLot2 As String, ByRef vsEmplacementDeFin As String, ByRef vbFinSaisie As Boolean) As Boolean

        Dim sScan As String = ""
        Dim iRes As Integer = 0

        RES_bSaisieEmplacementDeFin = False
        vsEmplacementDeFin = ""

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not RES_bSaisieEmplacementDeFin And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            RES_AffichageSaisieEMPLACEMENTdeFIN(vsEmplacementDeDebut, vsArticle, vsArticleLibelle, vsLot, vsQuantite, vsReferenceLot2)

            'Demande de saisie
            sScan = go_IO.RFInput(vsEmplacementFinalParDefaut, 10, CHR_nCentrer(10), 9, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else

                        vsEmplacementDeFin = Trim(sScan)
                        If (API_RES_bRechercheEmplacement(vsEmplacementDeFin)) Then
                            RES_bSaisieEmplacementDeFin = True
                        End If

                    End If
                Else
                    vbFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "RES_bSaisieEmplacementDeFin")
            End If

        End While

    End Function

    'Affichage du titre pour saisie de l'Emplacement de début
    Private Sub RES_AffichageSaisieEMPLACEMENTdeDEBUT()
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 2, CHR_sCentrer("EMPLACEMENT DEBUT"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage des articles trouvés sur les ID de stock de l'emplacement de début
    Private Sub RES_AffichageArticleIDStock(ByVal vsEmplacementDeDebut As String, ByVal vsTablo_Article As Object)

        Dim nIndex As Short = 0
        Dim vsArticleLibelle As String = ""
        Dim vsArticleLibelle_1 As String = ""
        Dim vsArticleLibelle_2 As String = ""
        Dim vsArticleLibelle_3 As String = ""
        Dim vsArticleLibelle_4 As String = ""
        Dim vsArticleLibelle_5 As String = ""
        Dim vsArticleLibelle_6 As String = ""

        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 1, "EMPL. : " & vsEmplacementDeDebut, WirelessStudioOle.RFIOConstants.WLNORMAL)

            For nIndex = 0 To UBound(vsTablo_Article)

                If (nIndex = 0) Then
                    .RFPrint(0, 2, "1- ART.: " & vsTablo_Article(nIndex), WirelessStudioOle.RFIOConstants.WLNORMAL)
                End If

                If (nIndex = 1) Then
                    .RFPrint(0, 3, "2- ART.: " & vsTablo_Article(nIndex), WirelessStudioOle.RFIOConstants.WLNORMAL)
                End If

                If (nIndex = 2) Then
                    .RFPrint(0, 4, "3- ART.: " & vsTablo_Article(nIndex), WirelessStudioOle.RFIOConstants.WLNORMAL)
                End If

                If (nIndex = 3) Then
                    .RFPrint(0, 5, "4- ART.: " & vsTablo_Article(nIndex), WirelessStudioOle.RFIOConstants.WLNORMAL)
                End If

                If (nIndex = 4) Then
                    .RFPrint(0, 6, "5- ART.: " & vsTablo_Article(nIndex), WirelessStudioOle.RFIOConstants.WLNORMAL)
                End If

                If (nIndex = 5) Then
                    .RFPrint(0, 7, "6- ART.: " & vsTablo_Article(nIndex), WirelessStudioOle.RFIOConstants.WLNORMAL)
                End If

            Next

            .RFPrint(0, 9, "N° ARTICLE", WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage des lots de l'article trouvés sur les ID de stock de l'emplacement de début
    Private Sub RES_AffichageLotIDStock(ByVal vsTypeEmplacementDeDebut As String, ByVal vsEmplacementDeDebut As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByVal vsLot As String, ByVal vsReferenceLot2 As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 1, "EMPL.: " & vsEmplacementDeDebut, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 2, "ART. : " & vsArticle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, vsArticleLibelle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 4, "LOT  : " & vsLot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            If (vsTypeEmplacementDeDebut = gTab_Configuration.sSLPT_Bord_De_Chaine_Normal) Then
                .RFPrint(0, 5, "LOT ORIGINE:" & vsReferenceLot2, WirelessStudioOle.RFIOConstants.WLNORMAL)
            End If
            .RFPrint(0, 8, CHR_sCentrer("LOT"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage de la demande de saisie de lot de l'article trouvé sur les ID de stock de l'emplacement de début
    Private Sub RES_AffichageLotASaisir(ByVal vsEmplacementDeDebut As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 1, "EMPL.: " & vsEmplacementDeDebut, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 2, "ART. : " & vsArticle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, vsArticleLibelle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 5, CHR_sCentrer("LOT D'ORIGINE"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage ID Stock pour saisie de la quantité
    Private Sub RES_AffichageIDStock(ByVal vsEmplacementDeDebut As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByVal vsLot As String, ByVal vsReferenceLot2 As String, ByVal vsQuantite As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 1, "EMPL.: " & vsEmplacementDeDebut, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 2, "ART. : " & vsArticle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, vsArticleLibelle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 4, "LOT  : " & vsLot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 5, "LOT ORIGINE:" & vsReferenceLot2, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 6, "QTE  : " & vsQuantite, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 8, CHR_sCentrer("QUANTITE"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage du titre pour saisie de l'Emplacement de fin
    Private Sub RES_AffichageSaisieEMPLACEMENTdeFIN(ByVal vsEmplacementDeDebut As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByVal vsLot As String,
                                                    ByVal vsQuantite As String, ByVal vsReferenceLot2 As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 1, "EMPL.: " & vsEmplacementDeDebut, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 2, "ART. : " & vsArticle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, vsArticleLibelle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 4, "LOT  : " & vsLot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 5, "LOT ORIGINE:" & vsReferenceLot2, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 6, "QTE  : " & vsQuantite, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 8, CHR_sCentrer("EMPLACEMENT FIN"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

End Module