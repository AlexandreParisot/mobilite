Option Strict Off
Option Explicit On
Module modEcranABC
    Dim msTitre As String

    'Entrée de l'option ABC - Approvisionnement bord de chaine
    Public Sub EcranABC(ByRef vsTitre As String)
        Dim bFinSaisie As Boolean = False
        Dim sEmplacementDeDebut As String = ""
        Dim sEmplacementDeFin As String = ""
        Dim sEAN128Saisi As String = ""
        Dim sNumOrdre As String = ""
        Dim sArticle As String = ""
        Dim sArticleLibelle As String = ""
        Dim sLot As String = ""
        Dim sQuantite As String = ""
        Dim sDateCourante As String = ""
        Dim sHeureCourante As String = ""
        Dim bSuiteTraitementOK As Boolean = False
        Dim sStatutIDStock As String = ""
        Dim sQuantiteAffectee As String = ""
        Dim sDatePeremption As String = ""
        Dim sTypeEmplacementDeFin As String = ""
        Dim sNouveauLot As String = ""
        Dim sQuantiteDejaAmenee As String = "0"
        Dim sDepot As String = ""
        Dim sDepotLibelle As String = ""


        msTitre = vsTitre

        If ABC_bSaisieDepot(sDepot, sDepotLibelle) Then

            If ABC_bSaisieEmplacementDeDebut(sDepot, sEmplacementDeDebut) Then

                While Not bFinSaisie

                    sEAN128Saisi = ""
                    ABC_bChoixScanOuSaisieEAN128(sDepot, sEmplacementDeDebut, sArticle, sQuantiteDejaAmenee, sEAN128Saisi, bFinSaisie)

                    sArticle = ""
                    sLot = ""
                    sQuantite = ""
                    sQuantiteDejaAmenee = "0"
                    If (Not bFinSaisie) Then

                        ' ====================
                        'Le CAB a été scanné ou saisi
                        If (sEAN128Saisi <> "") Then
                            If (CHR_bRecupInfoEAN128(sEAN128Saisi, sNumOrdre, sArticle, sLot, sQuantite)) Then

                                If (API_ABC_bControleArticleLot(sNumOrdre, sArticle, sLot)) Then

                                    If (API_ABC_bControleQuantiteArticleLot(sArticle, sLot, sQuantite, sDepot, sEmplacementDeDebut, sStatutIDStock, sQuantiteAffectee, sDatePeremption)) Then

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
                            If (ABC_bSaisieArticle(sDepot, sEmplacementDeDebut, sArticle, sArticleLibelle, bFinSaisie)) Then

                                If (ABC_bSaisieLot(sDepot, sEmplacementDeDebut, sArticle, sArticleLibelle, sLot, sNumOrdre, bFinSaisie)) Then

                                    If (API_ABC_bControleQuantiteArticleLot(sArticle, sLot, "0", sDepot, sEmplacementDeDebut, sStatutIDStock, sQuantiteAffectee, sDatePeremption)) Then

                                        bSuiteTraitementOK = True

                                    End If

                                End If

                            End If


                        End If
                    End If


                    ' ====================
                    'Suite du traitement
                    If (Not bFinSaisie And bSuiteTraitementOK = True) Then
                        'A ce stade on a recueilli les info Article/Lot (et la quantité si on a scannné le CAB) et on les a vérifié

                        If (sStatutIDStock = "2") Then

                            API_ABC_bControleDateLot(sDepot, sArticle, sLot, sStatutIDStock, sDatePeremption)

                            If (ABC_bSaisieQuantite(sDepot, sEmplacementDeDebut, sArticle, sArticleLibelle, sLot, sNumOrdre, sQuantite, sStatutIDStock, sQuantiteAffectee, sDatePeremption, bFinSaisie)) Then

                                If (ABC_bSaisieEmplacementDeFin(sDepot, sEmplacementDeDebut, sArticle, sArticleLibelle, sLot, sNumOrdre, sQuantite, sEmplacementDeFin, sTypeEmplacementDeFin, bFinSaisie)) Then

                                    API_bRecupDateHeure(sDateCourante, sHeureCourante)

                                    If (sTypeEmplacementDeFin = gTab_Configuration.sSLPT_Bord_De_Chaine_Normal) Then

                                        sNouveauLot = Trim(sDateCourante) & Trim(sHeureCourante)
                                        API_ABC_bReclassIdStock(sDepot, sEmplacementDeDebut, sArticle, sLot, sQuantite, sNouveauLot, sStatutIDStock)
                                        sLot = sNouveauLot

                                    ElseIf (sTypeEmplacementDeFin = gTab_Configuration.sSLPT_Bord_De_Chaine_Bib) Then

                                        sNouveauLot = Trim(sDateCourante)
                                        API_ABC_bReclassIdStock(sDepot, sEmplacementDeDebut, sArticle, sLot, sQuantite, sNouveauLot, sStatutIDStock)
                                        sLot = sNouveauLot

                                    End If

                                    If (API_ABC_bTransfertIdStock(sDepot, sEmplacementDeDebut, sArticle, sLot, sQuantite, sEmplacementDeFin)) Then
                                        API_ABC_bGestionQuantiteCumulee(sDepot, sArticle, sQuantite, sQuantiteDejaAmenee)
                                    End If

                                End If

                            End If

                        Else
                            MSG_AfficheErreur(giERR_ID_STOCK_STATUT_INVALIDE, sStatutIDStock, "2")
                        End If

                    End If

                End While

            End If

        End If

    End Sub

    'Saisie de l'emplacement
    Private Function ABC_bSaisieEmplacementDeDebut(ByVal vsDepot As String, ByRef vsEmplacementDeDebut As String) As Boolean

        ABC_bSaisieEmplacementDeDebut = False
        Dim bFinSaisie As Boolean = False
        Dim sScan As String = ""
        Dim iRes As Integer = 0

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not ABC_bSaisieEmplacementDeDebut And Not gbErreurCommunication And Not bFinSaisie

            'Affichage
            ABC_AffichageSaisieEMPLACEMENTdeDEBUT(vsDepot)

            'Demande de saisie
            sScan = go_IO.RFInput("", 10, CHR_nCentrer(10), 5, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        bFinSaisie = True
                    Else

                        vsEmplacementDeDebut = Trim(sScan)
                        If (API_ABC_bRechercheEmplacement(vsDepot, vsEmplacementDeDebut)) Then
                            ABC_bSaisieEmplacementDeDebut = True
                        End If

                    End If
                Else
                    bFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "ABC_bSaisieEmplacementDeDebut")
            End If

        End While

    End Function

    'Choix, soit Scan ou Saisie du code EAN128
    Private Function ABC_bChoixScanOuSaisieEAN128(ByVal vsDepot As String, ByVal vsEmplacement As String, ByVal vsArticle As String, ByVal vsQuantiteDejaAmenee As String, ByRef vsEAN128Saisi As String, ByRef vbFinSaisie As Boolean) As Boolean

        ABC_bChoixScanOuSaisieEAN128 = False
        Dim sScan As String = ""
        Dim iRes As Integer = 0
        vsEAN128Saisi = ""

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not ABC_bChoixScanOuSaisieEAN128 And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            ABC_AffichageChoix(vsDepot, vsEmplacement, vsArticle, vsQuantiteDejaAmenee)

            'Demande de saisie
            sScan = go_IO.RFInput("", 70, CHR_nCentrer(70), 8, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT_SAISIE, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else
                        If sScan = gCST_TOUCHE_SAISIE_F5 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                            ABC_bChoixScanOuSaisieEAN128 = True
                        Else
                            vsEAN128Saisi = sScan
                            ABC_bChoixScanOuSaisieEAN128 = True
                        End If
                    End If
                Else
                    vbFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "ABC_bChoixSsanOuSaisieEAN128")
            End If
        End While

    End Function

    'Saisie de l'article
    Private Function ABC_bSaisieArticle(ByVal vsDepot As String, ByVal vsEmplacementDeDebut As String, ByRef vsArticle As String, ByRef vsArticleLibelle As String, ByRef vbFinSaisie As Boolean) As Boolean

        ABC_bSaisieArticle = False
        Dim sScan As String = ""
        Dim iRes As Integer = 0

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not ABC_bSaisieArticle And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            ABC_AffichageSaisieARTICLE(vsDepot, vsEmplacementDeDebut)

            'Demande de saisie
            sScan = go_IO.RFInput("", 15, CHR_nCentrer(15), 6, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else

                        vsArticle = Trim(sScan)
                        If (API_ABC_bRechercheArticle(vsArticle, vsArticleLibelle)) Then
                            ABC_bSaisieArticle = True
                        End If

                    End If
                Else
                    vbFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "ABC_bSaisieArticle")
            End If

        End While

    End Function

    'Saisie du lot
    Private Function ABC_bSaisieLot(ByVal vsDepot As String, ByVal vsEmplacementdeDebut As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByRef vsLot As String, ByRef vsNumOrdre As String, ByRef vbFinSaisie As Boolean) As Boolean

        ABC_bSaisieLot = False
        Dim sScan As String = ""
        Dim iRes As Integer = 0
        vsLot = ""
        vsNumOrdre = ""

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not ABC_bSaisieLot And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            ABC_AffichageSaisieLOT(vsDepot, vsEmplacementdeDebut, vsArticle, vsArticleLibelle)

            'Demande de saisie
            sScan = go_IO.RFInput("", 20, CHR_nCentrer(20), 9, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else

                        vsLot = Trim(sScan)
                        If (API_ABC_bControleArticleLot(vsNumOrdre, vsArticle, vsLot)) Then
                            ABC_bSaisieLot = True
                        End If

                    End If
                Else
                    vbFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "ABC_bSaisieLot")
            End If

        End While

    End Function

    'Saisie de la quantité
    Private Function ABC_bSaisieQuantite(ByVal vsDepot As String, ByVal vsEmplacementDeDebut As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByVal vsLot As String, ByVal vsNumOrdre As String,
                                         ByRef vsQuantite As String, ByRef vsStatutIdStock As String, ByRef vsQuantiteAffectee As String, ByRef vsDatePeremption As String,
                                         ByRef vbFinSaisie As Boolean) As Boolean

        ABC_bSaisieQuantite = False
        Dim sScan As String = ""
        Dim iRes As Integer = 0
        Dim nQuantite As Long = 0
        Dim sQuantiteDejaAmenee As String = ""

        vsStatutIdStock = ""
        vsQuantiteAffectee = ""
        vsDatePeremption = ""

        API_ABC_bGestionQuantiteCumulee(vsDepot, vsArticle, "0", sQuantiteDejaAmenee)

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not ABC_bSaisieQuantite And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            ABC_AffichageSaisieQUANTITE(vsDepot, vsEmplacementDeDebut, vsArticle, vsArticleLibelle, vsLot, sQuantiteDejaAmenee)

            'Demande de saisie
            sScan = go_IO.RFInput(vsQuantite, 10, CHR_nCentrer(10), 11, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
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
                                If (API_ABC_bControleQuantiteArticleLot(vsArticle, vsLot, vsQuantite, vsDepot, vsEmplacementDeDebut, vsStatutIdStock, vsQuantiteAffectee, vsDatePeremption)) Then
                                    ABC_bSaisieQuantite = True
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
                WRL_GestionErreurPDT(iRes, "ABC_bSaisieQuantite")
            End If

        End While

    End Function

    'Saisie de l'emplacement de fin
    Private Function ABC_bSaisieEmplacementDeFin(ByVal vsDepot As String, ByVal vsEmplacementDeDebut As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByVal vsLot As String, ByVal vsNumOrdre As String,
                                                 ByVal vsQuantite As String, ByRef vsEmplacementDeFin As String, ByRef vsTypeEmplacementDeFin As String, ByRef vbFinSaisie As Boolean) As Boolean

        Dim sScan As String = ""
        Dim iRes As Integer = 0

        ABC_bSaisieEmplacementDeFin = False
        vsEmplacementDeFin = ""
        vsTypeEmplacementDeFin = ""

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not ABC_bSaisieEmplacementDeFin And Not gbErreurCommunication And Not vbFinSaisie

            'Affichage
            ABC_AffichageSaisieEMPLACEMENTdeFIN(vsDepot, vsEmplacementDeDebut, vsArticle, vsArticleLibelle, vsLot, vsQuantite)

            'Demande de saisie
            sScan = go_IO.RFInput("", 10, CHR_nCentrer(10), 11, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        vbFinSaisie = True
                    Else

                        vsEmplacementDeFin = Trim(sScan)
                        If (API_ABC_bRechercheEmplacement(vsDepot, vsEmplacementDeFin, vsTypeEmplacementDeFin)) Then
                            If (vsTypeEmplacementDeFin <> gTab_Configuration.sSLPT_Bord_De_Chaine_Normal And vsTypeEmplacementDeFin <> gTab_Configuration.sSLPT_Bord_De_Chaine_Bib) Then
                                MSG_AfficheErreur(giERR_TYPE_EMPLACEMENT_INVALIDE)
                            Else
                                ABC_bSaisieEmplacementDeFin = True
                            End If

                        End If

                    End If
                Else
                    vbFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "ABC_bSaisieEmplacementDeFin")
            End If

        End While

    End Function

    'Affichage du titre pour saisie de l'Emplacement de début
    Private Sub ABC_AffichageSaisieEMPLACEMENTdeDEBUT(ByVal vsDepot As String)
        With go_IO

            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 3, "DEPOT : " & vsDepot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 5, "EMPL. : ", WirelessStudioOle.RFIOConstants.WLNORMAL)

            '.RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            '.RFPrint(0, 2, CHR_sCentrer("EMPLACEMENT DEBUT"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage de l'écran choix de la saisie (Scan CAB ou saisie manuelle)
    Private Sub ABC_AffichageChoix(ByVal vsDepot As String, ByVal vsEmplacementDeDebut As String, ByVal vsArticle As String, ByVal vsQuantiteDejaAmenee As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 1, "DEPOT : " & vsDepot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 2, "EMPL. : " & vsEmplacementDeDebut, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 4, "ART.  : " & vsArticle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 5, "QTE   : " & vsQuantiteDejaAmenee, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 7, CHR_sCentrer("CAB PALETTE"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage du titre pour saisie de l'article
    Private Sub ABC_AffichageSaisieARTICLE(ByVal vsDepot As String, ByVal vsEmplacementDeDebut As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 2, "DEPOT : " & vsDepot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, "EMPL. : " & vsEmplacementDeDebut, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 5, CHR_sCentrer("ARTICLE"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage du titre pour saisie du lot
    Private Sub ABC_AffichageSaisieLOT(ByVal vsDepot As String, ByVal vsEmplacementDeDebut As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 2, "DEPOT : " & vsDepot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, "EMPL. : " & vsEmplacementDeDebut, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 5, "ART. : " & vsArticle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 6, vsArticleLibelle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 8, CHR_sCentrer("LOT"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage du titre pour saisie de la quantité
    Private Sub ABC_AffichageSaisieQUANTITE(ByVal vsDepot As String, ByVal vsEmplacementDeDebut As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByVal vsLot As String, ByVal vsQuantiteDejaAmenee As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 2, "DEPOT : " & vsDepot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, "EMPL. : " & vsEmplacementDeDebut, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 5, "ART. : " & vsArticle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 6, vsArticleLibelle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 7, "LOT : " & vsLot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 8, "QTE AMENEE : " & vsQuantiteDejaAmenee, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 10, CHR_sCentrer("QUANTITE"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    'Affichage du titre pour indiquer qu'un lot plus ancien existe
    Public Sub ABC_AffichageLotPlusAncien(ByVal vsDepot As String, ByVal vsEmplacement As String, ByVal vsLot As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 2, "DEPOT : " & vsDepot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, "Le lot ci-dessous est plus ancien", WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 5, "EMPL. : " & vsEmplacement, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 6, "LOT : " & vsLot, WirelessStudioOle.RFIOConstants.WLNORMAL)

            .GetEventEx(gCST_sFICHIER_BOUTONS_OK)
        End With
    End Sub

    'Affichage du titre pour saisie de l'Emplacement de fin
    Private Sub ABC_AffichageSaisieEMPLACEMENTdeFIN(ByVal vsDepot As String, ByVal vsEmplacementDeDebut As String, ByVal vsArticle As String, ByVal vsArticleLibelle As String, ByVal vsLot As String, ByVal vsQuantite As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 2, "DEPOT : " & vsDepot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 3, "EMPL. : " & vsEmplacementDeDebut, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 5, "ART. : " & vsArticle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 6, vsArticleLibelle, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 7, "LOT : " & vsLot, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 8, "QTE : " & vsQuantite, WirelessStudioOle.RFIOConstants.WLNORMAL)
            .RFPrint(0, 10, CHR_sCentrer("EMPLACEMENT FIN"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub

    Private Function ABC_bSaisieDepot(ByRef vsDepot As String, ByRef vsDepotLibelle As String) As Boolean

        ABC_bSaisieDepot = False
        Dim bFinSaisie As Boolean = False
        Dim sScan As String = ""
        Dim iRes As Integer = 0

        While sScan <> Chr(gCST_TOUCHE_ECHAP) And Not ABC_bSaisieDepot And Not gbErreurCommunication And Not bFinSaisie

            'Affichage
            ABC_AffichageSaisieDEPOT(vsDepot)

            'Demande de saisie
            sScan = go_IO.RFInput(gTab_Configuration.sDepot, 3, CHR_nCentrer(3), 4, gCST_sFICHIER_CODE_BARRE & "|" & gCST_sFICHIER_BOUTONS_OK_CLR_QUIT, WirelessStudioOle.RFIOConstants.WLKEEPKEYSTT, WirelessStudioOle.RFIOConstants.WLTOUPPER + WirelessStudioOle.RFIOConstants.WLNO_RETURN_BKSP + WirelessStudioOle.RFIOConstants.WLFORCE_ENTRY)
            iRes = go_IO.RFGetLastError() ' Gestion de l'erreur
            If iRes = WirelessStudioOle.RFErrorConstants.WLNOERROR Then
                If sScan <> Chr(gCST_TOUCHE_ECHAP) Then
                    If sScan = gCST_TOUCHE_QUITTER_F3 And go_IO.LastInputType = WirelessStudioOle.RFIOConstants.WLCOMMANDTYPE Then
                        bFinSaisie = True
                    Else

                        vsDepot = Trim(sScan)
                        If (API_EEP_bRechercheDepot(vsDepot, vsDepotLibelle)) Then
                            ABC_bSaisieDepot = True
                        End If

                    End If
                Else
                    bFinSaisie = True
                End If
            Else
                WRL_GestionErreurPDT(iRes, "ABC_bSaisieDepot")
            End If

        End While

    End Function

    'Affichage du titre pour saisie du dépôt
    Private Sub ABC_AffichageSaisieDEPOT(ByVal vsDepot As String)
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & msTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 2, CHR_sCentrer("DEPOT"), WirelessStudioOle.RFIOConstants.WLNORMAL)
        End With
    End Sub


End Module