Option Strict Off
Option Explicit On
Module modEcranABC_API

    'Recherche de l'emplacement
    'API=MMS010MI
    'Fonction=GetLocation
    Public Function API_ABC_bRechercheEmplacement(ByVal vsEmplacement As String, Optional ByRef vsTypeEmplacement As String = "") As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""

        API_ABC_bRechercheEmplacement = False
        vsTypeEmplacement = ""

        If API_bConnexionAPI("MMS010MI") Then

            'Construction de la fonction et de ses paramètAPI pour l'appel API
            sParam = CHR_sAjoutEspace("GetLocation", 15) &
                     CHR_sAjoutEspace(gTab_Configuration.sSociete, 3) &
                     CHR_sAjoutEspace(gTab_Configuration.sDepot, 3) &
                     CHR_sAjoutEspace(vsEmplacement, 10)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then
                    vsTypeEmplacement = Trim(Mid(sResultat, 179, 2))
                    API_ABC_bRechercheEmplacement = True
                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

    'Recherche Article/lot
    'API=MMS235MI
    'Fonction=GetItmLot
    Public Function API_ABC_bControleArticleLot(ByRef vsNumOrdre As String, ByVal vsArticle As String, ByVal vsLot As String) As Boolean

        API_ABC_bControleArticleLot = False
        Dim sParam As String = ""
        Dim sResultat As String = ""

        If API_bConnexionAPI("MMS235MI") Then

            'Construction de la fonction et de ses paramètAPI pour l'appel API
            sParam = CHR_sAjoutEspace("GetItmLot", 15) &
                     CHR_sAjoutEspace(vsArticle, 15) &
                     CHR_sAjoutEspace(vsLot, 20)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then

                    If (vsNumOrdre = "" Or (vsNumOrdre <> "" And vsNumOrdre = Trim(Mid(sResultat, 60, 10)))) Then
                        vsNumOrdre = Trim(Mid(sResultat, 60, 10))
                        API_ABC_bControleArticleLot = True
                    Else
                        MSG_AfficheErreur(giERR_CAB_NUM_ORDRE_INVALIDE)
                    End If

                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

    'Recherche de l'id de stock Article/lot
    'API=MMS060MI
    'Fonction=LstLot
    Public Function API_ABC_bControleQuantiteArticleLot(ByVal vsArticle As String, ByVal vsLot As String, ByVal vsQuantite As String, Optional ByVal vsEmplacement As String = "",
                                                        Optional ByRef vsStatutIDStock As String = "", Optional ByRef vsQuantiteAffectee As String = "", Optional ByRef vsDatePeremption As String = "") As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""
        Dim nListe As Long
        Dim nQuantite As Long = 0
        Dim nQuantiteStock As Long = 0

        API_ABC_bControleQuantiteArticleLot = False
        vsStatutIDStock = ""
        vsQuantiteAffectee = ""
        vsDatePeremption = ""

        If (IsNumeric(CHR_TransformeSeparateurPourNumerique(Trim(vsQuantite)))) Then
            nQuantite = (CHR_TransformeSeparateurPourNumerique(Trim(vsQuantite)))
        Else
            MSG_AfficheErreur(giERR_FORMAT_NUMERIC)
            Exit Function
        End If

        If API_bConnexionAPI("MMS060MI") Then

            'Construction de la fonction et de ses paramètAPI pour l'appel API
            sParam = CHR_sAjoutEspace("LstLot", 15) &
                     CHR_sAjoutEspace(gTab_Configuration.sDepot, 3) &
                     CHR_sAjoutEspace(vsLot, 20) &
                     CHR_sAjoutEspace(vsArticle, 15) &
                     CHR_sAjoutEspace(vsEmplacement, 10)

            If API_bTraitementAPI(sParam, sResultat, True) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then

                    While Mid(sResultat, 1, 3) = "REP"

                        If (IsNumeric(CHR_TransformeSeparateurPourNumerique(Trim(Mid(sResultat, 230, 11))))) Then
                            nQuantiteStock = nQuantiteStock + CHR_TransformeSeparateurPourNumerique(Trim(Mid(sResultat, 230, 11)))
                        End If

                        vsStatutIDStock = Trim(Mid(sResultat, 213, 1))
                        vsQuantiteAffectee = Trim(Mid(sResultat, 253, 11))
                        vsDatePeremption = Trim(Mid(sResultat, 222, 8))

                        API_RetourneResultatSuivantPourREP(sResultat, nListe)
                    End While

                    If (nQuantite > nQuantiteStock) Then
                        MSG_AfficheErreur(giERR_QTE_INVALIDE)
                    Else
                        API_ABC_bControleQuantiteArticleLot = True
                    End If

                Else
                    MSG_AfficheErreur(giERR_ID_STOCK_INVALIDE)
                End If
            End If

        End If

    End Function

    'Recherche de l'article
    'API=MMS200MI
    'Fonction=Get
    Public Function API_ABC_bRechercheArticle(ByVal vsArticle As String, ByRef vsArticleLibelle As String) As Boolean

        API_ABC_bRechercheArticle = False
        Dim sParam As String = ""
        Dim sResultat As String = ""
        vsArticleLibelle = ""

        If API_bConnexionAPI("MMS200MI") Then

            'Construction de la fonction et de ses paramètAPI pour l'appel API
            sParam = CHR_sAjoutEspace("Get", 15) &
                     CHR_sAjoutEspace(vsArticle, 15)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then
                    vsArticleLibelle = Trim(Mid(sResultat, 33, 30))
                    API_ABC_bRechercheArticle = True
                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

    'Recherche des id de stock Article/lot pour contrôle de la date de péremtion
    'API=MMS060MI
    'Fonction=LstBalID
    Public Function API_ABC_bControleDateLot(ByVal vsArticle As String, ByVal vsLot As String, ByVal vsStatutIDStock As String, ByVal vsDatePeremption As String) As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""
        Dim nListe As Long
        Dim sTemp_DatePeremption As String = ""
        Dim sTemp_Emplacement As String = ""
        Dim sTemp_Lot As String = ""

        API_ABC_bControleDateLot = False

        If API_bConnexionAPI("MMS060MI") Then

            'Construction de la fonction et de ses paramètAPI pour l'appel API
            sParam = CHR_sAjoutEspace("LstBalID", 15) &
                     CHR_sAjoutEspace(gTab_Configuration.sDepot, 3) &
                     CHR_sAjoutEspace(vsArticle, 15) &
                     CHR_sAjoutEspace("MS_OK", 10) &
                     CHR_sAjoutEspace("", 20) &
                     CHR_sAjoutEspace("", 35) &
                     CHR_sAjoutEspace("2", 1) &
                     CHR_sAjoutEspace("", 33) &
                     CHR_sAjoutEspace("1", 1)


            If API_bTraitementAPI(sParam, sResultat, True) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then

                    While Mid(sResultat, 1, 3) = "REP"

                        If (Trim(Mid(sResultat, 222, 8)) < vsDatePeremption) Then
                            sTemp_DatePeremption = Trim(Mid(sResultat, 222, 8))
                            sTemp_Emplacement = Trim(Mid(sResultat, 97, 10))
                            sTemp_Lot = Trim(Mid(sResultat, 107, 20))
                        End If

                        API_RetourneResultatSuivantPourREP(sResultat, nListe)
                    End While

                    If (sTemp_DatePeremption <> "") Then
                        ABC_AffichageLotPlusAncien(sTemp_Emplacement, sTemp_Lot)
                        API_ABC_bControleDateLot = True
                    Else
                        API_ABC_bControleDateLot = True
                    End If

                Else
                    MSG_AfficheErreur(giERR_ID_STOCK_INVALIDE)
                End If
            End If

        End If

    End Function

    'Gestion de la quantité cumulée qui a été amenée sur le bord de chaine depuis le changement de référence article.
    'Cette gestion se fait par adAPIse IP du terminal utilisé
    'API=CUSEXTMI
    'Fonction=xxxFieldValue
    Public Function API_ABC_bGestionQuantiteCumulee(ByVal vsArticle As String, ByVal vsQuantiteDeposee As String, ByRef vsQuantiteDejaAmenee As String) As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""
        Dim bChangeArticle As Boolean = False
        Dim bCreate As Boolean = False
        Dim bCumulQuantite As Boolean = False
        Dim lCumulQuantite As Long = 0
        Dim lQuantite As Long = 0
        Dim lQuantiteDeposee As Long = 0
        Dim lQuantiteDejaAmenee As Long = 0

        API_ABC_bGestionQuantiteCumulee = False
        vsQuantiteDejaAmenee = "0"


        If API_bConnexionAPI("CUSEXTMI") Then

            'Construction de la fonction et de ses paramètAPI pour l'appel API
            sParam = CHR_sAjoutEspace("GetFieldValue", 15) &
                     CHR_sAjoutEspace("RADIO_WS", 10) &
                     CHR_sAjoutEspace(gTab_Configuration.sSociete, 30) &
                     CHR_sAjoutEspace("SCENARIO_API", 30) &
                     CHR_sAjoutEspace(go_TRM.TerminalID, 30) &
                     CHR_sAjoutEspace("CUMULE_QTE", 30)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then

                    If (Trim(Mid(sResultat, 266, 30)) = Trim(vsArticle)) Then
                        If (IsNumeric(CHR_TransformeSeparateurPourNumerique(Trim(Mid(sResultat, 566, 21))))) Then
                            lQuantite = (CHR_TransformeSeparateurPourNumerique(Trim(Mid(sResultat, 566, 21))))
                        End If
                        vsQuantiteDejaAmenee = lQuantite
                        API_ABC_bGestionQuantiteCumulee = True

                        If (vsQuantiteDeposee <> "0") Then
                            bCumulQuantite = True
                            API_ABC_bGestionQuantiteCumulee = False
                        End If

                    Else
                        bChangeArticle = True
                    End If

                Else
                    bCreate = True
                End If
            End If

        End If

        ' ===============================
        If (bChangeArticle = True) Then
            'MAJ de l'enregistrement dans CUGEX1 pour changement d'article
            If API_bConnexionAPI("CUSEXTMI") Then

                'Construction de la fonction et de ses paramètAPI pour l'appel API
                sParam = CHR_sAjoutEspace("ChgFieldValue", 15) &
                         CHR_sAjoutEspace("RADIO_WS", 10) &
                         CHR_sAjoutEspace(gTab_Configuration.sSociete, 30) &
                         CHR_sAjoutEspace("SCENARIO_API", 30) &
                         CHR_sAjoutEspace(go_TRM.TerminalID, 30) &
                         CHR_sAjoutEspace("CUMULE_QTE", 30) &
                         CHR_sAjoutEspace("", 120) &
                         CHR_sAjoutEspace(vsArticle, 30) &
                         CHR_sAjoutEspace("", 270) &
                         CHR_sAjoutEspace(vsQuantiteDeposee, 21)

                If API_bTraitementAPI(sParam, sResultat) Then
                    If Mid(sResultat, 1, 3) <> "NOK" Then
                        vsQuantiteDejaAmenee = 0
                        API_ABC_bGestionQuantiteCumulee = True
                    Else
                        MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                    End If
                End If
            End If
        End If

        ' ===============================
        If (bCreate = True) Then
            'Création de l'enregistrement dans CUGEX1
            If API_bConnexionAPI("CUSEXTMI") Then

                'Construction de la fonction et de ses paramètAPI pour l'appel API
                sParam = CHR_sAjoutEspace("AddFieldValue", 15) &
                         CHR_sAjoutEspace("RADIO_WS", 10) &
                         CHR_sAjoutEspace(gTab_Configuration.sSociete, 30) &
                         CHR_sAjoutEspace("SCENARIO_API", 30) &
                         CHR_sAjoutEspace(go_TRM.TerminalID, 30) &
                         CHR_sAjoutEspace("CUMULE_QTE", 30) &
                        CHR_sAjoutEspace("", 120) &
                         CHR_sAjoutEspace(vsArticle, 30) &
                         CHR_sAjoutEspace("", 270) &
                         CHR_sAjoutEspace(vsQuantiteDeposee, 21)

                If API_bTraitementAPI(sParam, sResultat) Then
                    If Mid(sResultat, 1, 3) <> "NOK" Then
                        vsQuantiteDejaAmenee = 0
                        API_ABC_bGestionQuantiteCumulee = True
                    Else
                        MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                    End If
                End If
            End If
        End If

        ' ===============================
        If (bCumulQuantite = True) Then

            If (IsNumeric(CHR_TransformeSeparateurPourNumerique(Trim(vsQuantiteDeposee))) And IsNumeric(CHR_TransformeSeparateurPourNumerique(Trim(vsQuantiteDejaAmenee)))) Then
                lQuantiteDeposee = CHR_TransformeSeparateurPourNumerique(Trim(vsQuantiteDeposee))
                lQuantiteDejaAmenee = CHR_TransformeSeparateurPourNumerique(Trim(vsQuantiteDejaAmenee))
                lCumulQuantite = lQuantiteDeposee + lQuantiteDejaAmenee

                'MAJ de l'enregistrement dans CUGEX1 pour cumuler la quantité qui vient d'être déposée avec celle déjà amenée sur l'emplacement de bord de chaîne
                If API_bConnexionAPI("CUSEXTMI") Then

                    'Construction de la fonction et de ses paramètAPI pour l'appel API
                    sParam = CHR_sAjoutEspace("ChgFieldValue", 15) &
                             CHR_sAjoutEspace("RADIO_WS", 10) &
                             CHR_sAjoutEspace(gTab_Configuration.sSociete, 30) &
                             CHR_sAjoutEspace("SCENARIO_API", 30) &
                             CHR_sAjoutEspace(go_TRM.TerminalID, 30) &
                             CHR_sAjoutEspace("CUMULE_QTE", 30) &
                             CHR_sAjoutEspace("", 120) &
                             CHR_sAjoutEspace(vsArticle, 30) &
                             CHR_sAjoutEspace("", 270) &
                             CHR_sAjoutEspace(CStr(lCumulQuantite), 21, True)

                    If API_bTraitementAPI(sParam, sResultat) Then
                        If Mid(sResultat, 1, 3) <> "NOK" Then
                            vsQuantiteDejaAmenee = CStr(lCumulQuantite)
                            API_ABC_bGestionQuantiteCumulee = True
                        Else
                            MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                        End If
                    End If
                End If

            End If

        End If

    End Function

    'Reclassification de l'ID de stock
    'API=MMS850MI
    'Fonction=AddReclass
    Public Function API_ABC_bReclassIdStock(ByVal vsEmplacement As String, ByVal vsArticle As String, ByVal vsLot As String, ByVal vsQuantite As String,
                                            ByVal vsNouveauLot As String, ByVal vsStatutIDStock As String) As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""

        API_ABC_bReclassIdStock = False

        If API_bConnexionAPI("MMS850MI") Then

            'Construction de la fonction et de ses paramètAPI pour l'appel API
            sParam = CHR_sAjoutEspace("AddReclass", 15) &
                     CHR_sAjoutEspace("*EXE", 4) &
                     CHR_sAjoutEspace("", 42) &
                     CHR_sAjoutEspace("WIRELESS", 17) &
                     CHR_sAjoutEspace("WMS", 6) &
                     CHR_sAjoutEspace(gTab_Configuration.sDepot, 3) &
                     CHR_sAjoutEspace(vsEmplacement, 10) &
                     CHR_sAjoutEspace(vsArticle, 15) &
                     CHR_sAjoutEspace(vsLot, 20) &
                     CHR_sAjoutEspace("", 50) &
                     CHR_sAjoutEspace(vsLot, 20) &
                     CHR_sAjoutEspace("", 56) &
                     CHR_sAjoutEspace(vsQuantite, 17, True) &
                     CHR_sAjoutEspace("", 13) &
                     CHR_sAjoutEspace(vsArticle, 15) &
                     CHR_sAjoutEspace(vsNouveauLot, 20) &
                     CHR_sAjoutEspace("", 21) &
                     CHR_sAjoutEspace(vsStatutIDStock, 1)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then
                    API_ABC_bReclassIdStock = True
                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

    'Transfert du stock
    'API=MMS175MI
    'Fonction=Update
    Public Function API_ABC_bTransfertIdStock(ByVal vsEmplacementDeDebut As String, ByVal vsArticle As String, ByVal vsLot As String, ByVal vsQuantite As String,
                                            ByVal vsEmplacementDeFin As String) As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""

        API_ABC_bTransfertIdStock = False

        If API_bConnexionAPI("MMS175MI") Then

            'Construction de la fonction et de ses paramètAPI pour l'appel API
            sParam = CHR_sAjoutEspace("Update", 15) &
                     CHR_sAjoutEspace(gTab_Configuration.sSociete, 3) &
                     CHR_sAjoutEspace(gTab_Configuration.sDepot, 3) &
                     CHR_sAjoutEspace(vsArticle, 15) &
                     CHR_sAjoutEspace(vsEmplacementDeFin, 10) &
                     CHR_sAjoutEspace(vsQuantite, 11, True) &
                     CHR_sAjoutEspace(vsEmplacementDeDebut, 10) &
                     CHR_sAjoutEspace(vsLot, 20)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then
                    API_ABC_bTransfertIdStock = True
                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

End Module


