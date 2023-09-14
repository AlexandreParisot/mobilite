Option Strict Off
Option Explicit On
Module modEcranRES_API

    'Recherche de l'emplacement
    'API=MMS010MI
    'Fonction=GetLocation
    Public Function API_RES_bRechercheEmplacement(ByVal vsEmplacement As String, Optional ByRef vsTypeEmplacement As String = "") As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""

        API_RES_bRechercheEmplacement = False
        vsTypeEmplacement = ""

        If API_bConnexionAPI("MMS010MI") Then

            'Construction de la fonction et de ses paramètres pour l'appel API
            sParam = CHR_sAjoutEspace("GetLocation", 15) &
                     CHR_sAjoutEspace(gTab_Configuration.sSociete, 3) &
                     CHR_sAjoutEspace(gTab_Configuration.sDepot, 3) &
                     CHR_sAjoutEspace(vsEmplacement, 10)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then
                    vsTypeEmplacement = Trim(Mid(sResultat, 179, 2))
                    API_RES_bRechercheEmplacement = True
                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

    'Recherche Recherche des derniers ID de stock entrés sur l'emplacement par utilisateur
    'API=ZZZ002MI
    'Fonction=LstLastIDStock
    Public Function API_RES_bRechercheDernierIdStockSurEmplacement(ByVal vsEmplacement As String, ByVal vsArticle As String, ByVal vsLot As String, ByRef vsTablo_Article As Object,
                                                                   ByRef vsTablo_Lot As Object, ByRef vsTablo_BRE2 As Object, ByRef vsTablo_Qte As Object) As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""
        Dim nListe As Long = 0
        Dim nIndex As Short = 0
        Dim sTablo_Article(0)
        Dim sTablo_Lot(0)
        Dim sTablo_BRE2(0)
        Dim sTablo_Qte(0)

        Dim lQuantite As Long = 0

        API_RES_bRechercheDernierIdStockSurEmplacement = False

        If API_bConnexionAPI("ZZZ002MI") Then
            ReDim sTablo_Article(0)
            ReDim sTablo_Lot(0)
            ReDim sTablo_BRE2(0)
            ReDim sTablo_Qte(0)

            'Construction de la fonction et de ses paramètres pour l'appel API
            sParam = CHR_sAjoutEspace("LstLastIdStock", 15) &
                         CHR_sAjoutEspace(gTab_Configuration.sDepot, 3) &
                         CHR_sAjoutEspace(vsEmplacement, 10) &
                         CHR_sAjoutEspace(gTab_Configuration.sUtilisateur, 10) &
                         CHR_sAjoutEspace("", 8) &
                         CHR_sAjoutEspace("", 6) &
                         CHR_sAjoutEspace(vsArticle, 15) &
                         CHR_sAjoutEspace(vsLot, 20)

            If API_bTraitementAPI(sParam, sResultat, True) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then

                    While Mid(sResultat, 1, 3) = "REP"

                        ReDim Preserve sTablo_Article(nIndex)
                        ReDim Preserve sTablo_Lot(nIndex)
                        ReDim Preserve sTablo_BRE2(nIndex)
                        ReDim Preserve sTablo_Qte(nIndex)

                        sTablo_Article(nIndex) = Trim(Mid(sResultat, 16, 15))
                        sTablo_Lot(nIndex) = Trim(Mid(sResultat, 31, 20))
                        sTablo_BRE2(nIndex) = Trim(Mid(sResultat, 51, 20))

                        If (IsNumeric(CHR_TransformeSeparateurPourNumerique(Trim(Mid(sResultat, 71, 17))))) Then
                            lQuantite = (CHR_TransformeSeparateurPourNumerique(Trim(Mid(sResultat, 71, 17))))
                        End If
                        sTablo_Qte(nIndex) = lQuantite

                        nIndex = nIndex + 1
                        API_RetourneResultatSuivantPourREP(sResultat, nListe)
                    End While

                    'Si aucun ID de stock trouvé, on averti l'utilisateur 
                    If (nIndex = 0) Then
                        MSG_AfficheErreur(giERR_ID_STOCK_NON_TROUVE, sResultat)
                    Else
                        vsTablo_Article = sTablo_Article
                        vsTablo_Lot = sTablo_Lot
                        vsTablo_BRE2 = sTablo_BRE2
                        vsTablo_Qte = sTablo_Qte
                    End If

                Else
                    'Si aucun ID de stock trouvé, on averti l'utilisateur 
                    MSG_AfficheErreur(giERR_ID_STOCK_NON_TROUVE, sResultat)
                End If

            End If

        End If

    End Function

    'Recherche de l'id de stock Article/lot
    'API=MMS060MI
    'Fonction=LstLot
    Public Function API_RES_bControleQuantiteArticleLot(ByVal vsArticle As String, ByVal vsLot As String, ByVal vsQuantite As String, Optional ByVal vsEmplacement As String = "",
                                                        Optional ByRef vsStatutIDStock As String = "", Optional ByRef vsReferenceLot2 As String = "") As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""
        Dim nListe As Long
        Dim nQuantite As Long = 0
        Dim nQuantiteStock As Long = 0

        API_RES_bControleQuantiteArticleLot = False
        vsStatutIDStock = ""
        vsReferenceLot2 = ""

        If (IsNumeric(CHR_TransformeSeparateurPourNumerique(Trim(vsQuantite)))) Then
            nQuantite = (CHR_TransformeSeparateurPourNumerique(Trim(vsQuantite)))
        Else
            MSG_AfficheErreur(giERR_FORMAT_NUMERIC)
            Exit Function
        End If

        If API_bConnexionAPI("MMS060MI") Then

            'Construction de la fonction et de ses paramètres pour l'appel API
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
                        vsReferenceLot2 = Trim(Mid(sResultat, 341, 20))

                        API_RetourneResultatSuivantPourREP(sResultat, nListe)
                    End While

                    If (nQuantite > nQuantiteStock) Then
                        MSG_AfficheErreur(giERR_QTE_INVALIDE)
                    Else
                        API_RES_bControleQuantiteArticleLot = True
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
    Public Function API_RES_bRechercheArticle(ByVal vsArticle As String, ByRef vsArticleLibelle As String, ByRef vsArticleType As String) As Boolean

        API_RES_bRechercheArticle = False
        Dim sParam As String = ""
        Dim sResultat As String = ""
        vsArticleLibelle = ""
        vsArticleType = ""

        If API_bConnexionAPI("MMS200MI") Then

            'Construction de la fonction et de ses paramètres pour l'appel API
            sParam = CHR_sAjoutEspace("Get", 15) &
                     CHR_sAjoutEspace(vsArticle, 15)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then
                    vsArticleLibelle = Trim(Mid(sResultat, 33, 30))
                    vsArticleType = Trim(Mid(sResultat, 172, 3))
                    API_RES_bRechercheArticle = True
                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

    'Recherche de l'article/dépôt
    'API=MMS200MI
    'Fonction=GetItmWhsBasic
    Public Function API_RES_bRechercheArticleDepot(ByVal vsArticle As String, ByRef vsEmplacementArticleDepot As String) As Boolean


        Dim sParam As String = ""
        Dim sResultat As String = ""
        API_RES_bRechercheArticleDepot = False
        vsEmplacementArticleDepot = ""

        If API_bConnexionAPI("MMS200MI") Then

            'Construction de la fonction et de ses paramètres pour l'appel API
            sParam = CHR_sAjoutEspace("GetItmWhsBasic", 15) &
                     CHR_sAjoutEspace(gTab_Configuration.sSociete, 3) &
                     CHR_sAjoutEspace(gTab_Configuration.sDepot, 3) &
                     CHR_sAjoutEspace(vsArticle, 15)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then
                    vsEmplacementArticleDepot = Trim(Mid(sResultat, 176, 10))
                    API_RES_bRechercheArticleDepot = True
                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

    'Reclassification de l'ID de stock
    'API=MMS850MI
    'Fonction=AddReclass
    Public Function API_RES_bReclassIdStock(ByVal vsEmplacement As String, ByVal vsArticle As String, ByVal vsLot As String, ByVal vsQuantite As String,
                                            ByVal vsNouveauLot As String, ByVal vsStatutIDStock As String) As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""

        API_RES_bReclassIdStock = False

        If API_bConnexionAPI("MMS850MI") Then

            'Construction de la fonction et de ses paramètres pour l'appel API
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
                     CHR_sAjoutEspace("-", 20) &
                     CHR_sAjoutEspace("", 56) &
                     CHR_sAjoutEspace(vsQuantite, 17, True) &
                     CHR_sAjoutEspace("", 13) &
                     CHR_sAjoutEspace(vsArticle, 15) &
                     CHR_sAjoutEspace(vsNouveauLot, 20) &
                     CHR_sAjoutEspace("", 21) &
                     CHR_sAjoutEspace(vsStatutIDStock, 1)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then
                    API_RES_bReclassIdStock = True
                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

    'Transfert du stock
    'API=MMS175MI
    'Fonction=Update
    Public Function API_RES_bTransfertIdStock(ByVal vsEmplacementDeDebut As String, ByVal vsArticle As String, ByVal vsLot As String, ByVal vsQuantite As String,
                                            ByVal vsEmplacementDeFin As String) As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""

        API_RES_bTransfertIdStock = False

        If API_bConnexionAPI("MMS175MI") Then

            'Construction de la fonction et de ses paramètres pour l'appel API
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
                    API_RES_bTransfertIdStock = True
                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

    'Edition de l'étiquette palette
    'API=MMS060MI
    'Fonction=PrtPutAwayLbl
    Public Function API_RES_bEditionEtiquettePalette(ByVal vsEmplacement As String, ByVal vsNumOrdre As String, ByVal vsArticle As String, ByVal vsLot As String, ByVal vsQuantite As String) As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""

        API_RES_bEditionEtiquettePalette = False

        If API_bConnexionAPI("MMS060MI") Then

            'Construction de la fonction et de ses paramètres pour l'appel API
            sParam = CHR_sAjoutEspace("PrtPutAwayLbl", 15) &
                CHR_sAjoutEspace(gTab_Configuration.sDepot, 3) &
                     CHR_sAjoutEspace(vsArticle, 15) &
                     CHR_sAjoutEspace(vsEmplacement, 10) &
                     CHR_sAjoutEspace(vsLot, 20) &
                     CHR_sAjoutEspace("", 20) &
                     CHR_sAjoutEspace("", 10) &
                     CHR_sAjoutEspace("", 10) &
                     CHR_sAjoutEspace("", 3) &
                     CHR_sAjoutEspace(vsQuantite, 17)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then

                    API_RES_bEditionEtiquettePalette = True

                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

    'Recherche de la dernière édition de l'étiquette palette, puis on la ré-imprime
    'API=ZZZ002MI
    'Fonction=RtvLastPrintout
    Public Function API_RES_bImpression_Derniere_Etiquette() As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""

        API_RES_bImpression_Derniere_Etiquette = False

        If API_bConnexionAPI("ZZZ002MI") Then

            'Construction de la fonction et de ses paramètres pour l'appel API
            sParam = CHR_sAjoutEspace("RtvLastPrintout", 15) &
                     CHR_sAjoutEspace("SAV ", 4) &
                     CHR_sAjoutEspace("MWS450PF", 10) &
                     CHR_sAjoutEspace(gTab_Configuration.sUtilisateur, 10)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then
                    API_RES_bImpression_Derniere_Etiquette = True

                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

    'Recherche Article/lot
    'API=MMS235MI
    'Fonction=GetItmLot
    Public Function API_RES_bControleArticleLot(ByVal vsArticle As String, ByVal vsLot As String) As Boolean

        API_RES_bControleArticleLot = False
        Dim sParam As String = ""
        Dim sResultat As String = ""

        If API_bConnexionAPI("MMS235MI") Then

            'Construction de la fonction et de ses paramètAPI pour l'appel API
            sParam = CHR_sAjoutEspace("GetItmLot", 15) &
                     CHR_sAjoutEspace(vsArticle, 15) &
                     CHR_sAjoutEspace(vsLot, 20)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then

                    API_RES_bControleArticleLot = True

                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

End Module




