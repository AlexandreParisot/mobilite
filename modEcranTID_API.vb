Option Strict Off
Option Explicit On
Module modEcranTID_API

    'Recherche de la LP
    'API=MWS420MI
    'Fonction=LstPickHeader
    Public Function API_TID_bRechercheLP(ByVal vsIndexDeLivraison As String, ByVal vsNumLP As String) As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""

        API_TID_bRechercheLP = False

        If API_bConnexionAPI("MWS420MI") Then

            'Construction de la fonction et de ses paramètres pour l'appel API
            sParam = CHR_sAjoutEspace("LstPickHeader", 15) &
                     CHR_sAjoutEspace(vsIndexDeLivraison, 11) &
                     CHR_sAjoutEspace(vsNumLP, 3)

            If API_bTraitementAPI(sParam, sResultat, True) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then

                    If (Mid(sResultat, 1, 3) = "REP") Then
                        API_FermeConnexionPourREP()
                        API_TID_bRechercheLP = True
                    Else
                        MSG_AfficheErreur(giERR_LP_INVALIDE)
                    End If

                Else
                    MSG_AfficheErreur(giERR_LP_INVALIDE)
                End If
            End If

        End If

    End Function

    'Recherche des lignes de LP à préparer
    'API=MW422MI
    'Fonction=SelPickDetail
    Public Function API_TID_bRechercheLigneLP_A_Preparer(ByVal vsIndexDeLivraison As String, ByVal vsNumLP As String, ByRef vsArticle As String, ByRef vsArticleLibelle As String, ByRef vsQuantite As String,
                                                         ByRef vsNumLigneLP As String, ByRef vsSensLigneLP As String) As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""
        Dim nListe As Long
        Dim nQuantite As Long = 0
        Dim sNumLigneLP_Encours As String = ""
        Dim sNumLigneLP_Precedant As String = ""
        Dim sArticle_Precedant As String = ""
        Dim sArticleLibelle_Precedant As String = ""
        Dim sQuantite_Precedant As String = ""
        Dim bLigneTrouve As Boolean = False

        API_TID_bRechercheLigneLP_A_Preparer = False
        vsArticle = ""
        vsArticleLibelle = ""
        vsQuantite = ""

        If API_bConnexionAPI("MWS422MI") Then

            'Construction de la fonction et de ses paramètAPI pour l'appel API
            sParam = CHR_sAjoutEspace("SelPickDetail", 15) &
                     CHR_sAjoutEspace(vsIndexDeLivraison, 11) &
                     CHR_sAjoutEspace(vsNumLP, 3) &
                     CHR_sAjoutEspace("", 114) &
                     CHR_sAjoutEspace("40", 20)

            If API_bTraitementAPI(sParam, sResultat, True) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then

                    While Mid(sResultat, 1, 3) = "REP"

                        If (Trim(Mid(sResultat, 301, 1)) = "1") Then    '==> signifie uniquement les lignes en Soft allocation

                            sNumLigneLP_Encours = Trim(Mid(sResultat, 435, 11))

                            If (vsSensLigneLP = "") Then    '==> Signifie que l'utilisateur veut voir la 1 ère ligne de LP à préparer
                                vsNumLigneLP = sNumLigneLP_Encours
                                vsArticle = Trim(Mid(sResultat, 22, 15))
                                vsArticleLibelle = Trim(Mid(sResultat, 488, 30))
                                vsQuantite = Trim(Mid(sResultat, 141, 17))
                                API_FermeConnexionPourREP()
                                API_TID_bRechercheLigneLP_A_Preparer = True
                                Exit While

                            ElseIf (vsSensLigneLP = "+") Then   '==> Signifie que l'utilisateur veut voir la ligne de LP suivante à préparer

                                If (vsNumLigneLP = sNumLigneLP_Encours) Then
                                    bLigneTrouve = True
                                Else
                                    If (bLigneTrouve = True) Then
                                        vsNumLigneLP = sNumLigneLP_Encours
                                        vsArticle = Trim(Mid(sResultat, 22, 15))
                                        vsArticleLibelle = Trim(Mid(sResultat, 488, 30))
                                        vsQuantite = Trim(Mid(sResultat, 141, 17))
                                        API_FermeConnexionPourREP()
                                        API_TID_bRechercheLigneLP_A_Preparer = True
                                        Exit While
                                    End If
                                End If

                            ElseIf (vsSensLigneLP = "-") Then   '==> Signifie que l'utilisateur veut voir la ligne de LP précédente à préparer

                                If (vsNumLigneLP = sNumLigneLP_Encours) Then

                                    If (sNumLigneLP_Precedant = "") Then
                                        vsNumLigneLP = sNumLigneLP_Encours
                                        vsArticle = Trim(Mid(sResultat, 22, 15))
                                        vsArticleLibelle = Trim(Mid(sResultat, 488, 30))
                                        vsQuantite = Trim(Mid(sResultat, 141, 17))
                                    Else
                                        vsNumLigneLP = sNumLigneLP_Precedant
                                        vsArticle = sArticle_Precedant
                                        vsArticleLibelle = sArticleLibelle_Precedant
                                        vsQuantite = sQuantite_Precedant
                                    End If

                                    API_FermeConnexionPourREP()
                                    API_TID_bRechercheLigneLP_A_Preparer = True
                                    Exit While

                                Else
                                    sNumLigneLP_Precedant = sNumLigneLP_Encours
                                    sArticle_Precedant = Trim(Mid(sResultat, 22, 15))
                                    sArticleLibelle_Precedant = Trim(Mid(sResultat, 488, 30))
                                    sQuantite_Precedant = Trim(Mid(sResultat, 141, 17))
                                End If

                            End If

                        End If

                        API_RetourneResultatSuivantPourREP(sResultat, nListe)
                    End While

                    If (vsNumLigneLP = "" And vsSensLigneLP = "") Then
                        vsArticle = ""
                        vsArticleLibelle = ""
                        vsQuantite = ""
                        MSG_AfficheErreur(giERR_PLUS_DE_LIGNE_DE_LP)
                    End If

                Else
                    vsArticle = ""
                    vsArticleLibelle = ""
                    vsQuantite = ""
                    vsNumLigneLP = ""
                    MSG_AfficheErreur(giERR_PLUS_DE_LIGNE_DE_LP)
                End If
            End If

        End If

    End Function

    'Recherche d'info sur l'index de livraison
    'API=MWS410MI
    'Fonction=GetHead
    Public Function API_TID_bRtvInfoLP(ByVal vsIndexDeLivraison As String, ByVal vsNumLP As String, ByRef vsNumOD As String, ByRef vsDepotDebut As String, ByRef vsDepotFin As String) As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""
        API_TID_bRtvInfoLP = False
        vsNumOD = ""
        vsDepotDebut = ""
        vsDepotFin = ""

        If API_bConnexionAPI("MWS410MI") Then

            'Construction de la fonction et de ses paramètAPI pour l'appel API
            sParam = CHR_sAjoutEspace("GetHead", 15) &
                     CHR_sAjoutEspace(gTab_Configuration.sSociete, 3) &
                     CHR_sAjoutEspace(vsIndexDeLivraison, 11)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then

                    If (Trim(Mid(sResultat, 327, 1)) = "5") Then
                        vsNumOD = Trim(Mid(sResultat, 331, 10))
                        vsDepotDebut = Trim(Mid(sResultat, 328, 3))
                        vsDepotFin = Trim(Mid(sResultat, 669, 10))
                        API_TID_bRtvInfoLP = True
                    Else
                        MSG_AfficheErreur(giERR_LP_INVALIDE)
                    End If

                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

    'Recherche Article/Dépôt
    'API=MMS200MI
    'Fonction=GetItmWhsBasic
    Public Function API_TID_bControleArticleDepot(ByVal vsDepot As String, ByVal vsArticle As String, ByRef vsEmplacementArticleDepot As String) As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""

        API_TID_bControleArticleDepot = False
        vsEmplacementArticleDepot = ""

        If API_bConnexionAPI("MMS200MI") Then

            'Construction de la fonction et de ses paramètAPI pour l'appel API
            sParam = CHR_sAjoutEspace("GetItmWhsBasic", 15) &
                     CHR_sAjoutEspace(gTab_Configuration.sSociete, 3) &
                     CHR_sAjoutEspace(vsDepot, 3) &
                     CHR_sAjoutEspace(vsArticle, 15)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then
                    vsEmplacementArticleDepot = Trim(Mid(sResultat, 176, 10))
                    API_TID_bControleArticleDepot = True
                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

    'Recherche Article/lot
    'API=MMS235MI
    'Fonction=GetItmLot
    Public Function API_TID_bControleArticleLot(ByRef vsNumOrdre As String, ByVal vsArticle As String, ByVal vsLot As String) As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""
        API_TID_bControleArticleLot = False

        If API_bConnexionAPI("MMS235MI") Then

            'Construction de la fonction et de ses paramètAPI pour l'appel API
            sParam = CHR_sAjoutEspace("GetItmLot", 15) &
                     CHR_sAjoutEspace(vsArticle, 15) &
                     CHR_sAjoutEspace(vsLot, 20)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then

                    If (vsNumOrdre = "" Or (vsNumOrdre <> "" And vsNumOrdre = Trim(Mid(sResultat, 60, 10)))) Then
                        vsNumOrdre = Trim(Mid(sResultat, 60, 10))
                        API_TID_bControleArticleLot = True
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
    Public Function API_TID_bControleQuantiteArticleLot(ByVal vsDepot As String, ByVal vsArticle As String, ByVal vsLot As String, ByVal vsQuantite As String, Optional ByVal vsEmplacement As String = "",
                                                        Optional ByRef vsStatutIDStock As String = "", Optional ByRef vsQuantiteAffectee As String = "", Optional ByRef vsDatePeremption As String = "") As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""
        Dim nListe As Long
        Dim nQuantite As Long = 0
        Dim nQuantiteStock As Long = 0

        API_TID_bControleQuantiteArticleLot = False
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
                     CHR_sAjoutEspace(vsDepot, 3) &
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
                        API_TID_bControleQuantiteArticleLot = True
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
    Public Function API_TID_bRechercheArticle(ByVal vsArticle As String, ByRef vsArticleLibelle As String, Optional ByRef vsArticleType As String = "") As Boolean

        API_TID_bRechercheArticle = False
        Dim sParam As String = ""
        Dim sResultat As String = ""
        vsArticleLibelle = ""

        If API_bConnexionAPI("MMS200MI") Then

            'Construction de la fonction et de ses paramètres pour l'appel API
            sParam = CHR_sAjoutEspace("Get", 15) &
                     CHR_sAjoutEspace(vsArticle, 15)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then
                    vsArticleLibelle = Trim(Mid(sResultat, 33, 30))
                    vsArticleType = Trim(Mid(sResultat, 172, 3))
                    API_TID_bRechercheArticle = True
                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

    'Recherche des lignes de LP à préparer
    'API=MW422MI
    'Fonction=SelPickDetail
    Public Function API_TID_bRechercheLigneLP_Article(ByVal vsIndexDeLivraison As String, ByVal vsNumLP As String, ByVal vsArticle As String, ByRef vsNumLigneLP As String,
                                                      ByRef vsQuantiteRestanteAPreparer As String) As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""
        Dim nListe As Long

        API_TID_bRechercheLigneLP_Article = False
        vsNumLigneLP = ""
        vsQuantiteRestanteAPreparer = ""

        If API_bConnexionAPI("MWS422MI") Then

            'Construction de la fonction et de ses paramètAPI pour l'appel API
            sParam = CHR_sAjoutEspace("SelPickDetail", 15) &
                     CHR_sAjoutEspace(vsIndexDeLivraison, 11) &
                     CHR_sAjoutEspace(vsNumLP, 3) &
                     CHR_sAjoutEspace("", 38) &
                     CHR_sAjoutEspace(vsArticle, 15) &
                     CHR_sAjoutEspace("", 61) &
                     CHR_sAjoutEspace("40", 20)

            If API_bTraitementAPI(sParam, sResultat, True) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then

                    While Mid(sResultat, 1, 3) = "REP"

                        If (Trim(Mid(sResultat, 301, 1)) = "1") Then    '==> signifie uniquement les lignes en Soft allocation

                            vsNumLigneLP = Trim(Mid(sResultat, 435, 11))
                            vsQuantiteRestanteAPreparer = Trim(Mid(sResultat, 141, 17))
                            API_FermeConnexionPourREP()
                            API_TID_bRechercheLigneLP_Article = True
                            Exit While

                        End If

                        API_RetourneResultatSuivantPourREP(sResultat, nListe)
                    End While

                    If (vsNumLigneLP = "") Then
                        MSG_AfficheErreur(giERR_PLUS_DE_LIGNE_DE_LP)
                    End If

                Else
                    MSG_AfficheErreur(giERR_ARTICLE_INEXISTANT_DANS_LP)
                End If
            End If

        End If

    End Function

    'Validation de la ligne de LP
    'API=MHS850MI
    'Fonction=AddPickViaRepNo
    Public Function API_TID_bValideLigneLP(ByVal vsIndexDeLivraison As String, ByVal vsNumLP As String, ByVal vsDepot As String, ByVal vsEmplacement As String, ByVal vsArticle As String,
                                           ByVal vsLot As String, ByVal vsQuantite As String, ByVal vsNumLigneLP As String, ByVal vsQuantiteRestanteAPreparer As String) As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""
        Dim sMarqueDeFin As String = ""

        API_TID_bValideLigneLP = False

        If (vsQuantite = "0") Then
            sMarqueDeFin = "1"
        Else
            vsQuantiteRestanteAPreparer = ""
        End If

        If API_bConnexionAPI("MHS850MI") Then

            'Construction de la fonction et de ses paramètres pour l'appel API
            sParam = CHR_sAjoutEspace("AddPickViaRepNo", 15) &
                     CHR_sAjoutEspace("*EXE", 4) &
                     CHR_sAjoutEspace(gTab_Configuration.sSociete, 3) &
                     CHR_sAjoutEspace(vsDepot, 3) &
                     CHR_sAjoutEspace("", 56) &
                     CHR_sAjoutEspace("WIRELESS", 17) &
                     CHR_sAjoutEspace("WMS", 6) &
                     CHR_sAjoutEspace(vsNumLigneLP, 11) &
                     CHR_sAjoutEspace(vsQuantite, 17) &
                     CHR_sAjoutEspace(vsQuantiteRestanteAPreparer, 17) &
                     CHR_sAjoutEspace(vsEmplacement, 10) &
                     CHR_sAjoutEspace(vsLot, 20) &
                     CHR_sAjoutEspace("", 30) &
                     CHR_sAjoutEspace(sMarqueDeFin, 1) &
                     CHR_sAjoutEspace("", 136) &
                     CHR_sAjoutEspace("2", 2)


            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then
                    API_TID_bValideLigneLP = True
                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

    'Validation de la LP
    'API=MWS420MI
    'Fonction=Confirm
    Public Function API_TID_bValideLP(ByVal vsIndexDeLivraison As String, ByVal vsNumLP As String) As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""

        API_TID_bValideLP = False

        If API_bConnexionAPI("MWS420MI") Then

            'Construction de la fonction et de ses paramètres pour l'appel API
            sParam = CHR_sAjoutEspace("Confirm", 15) &
                     CHR_sAjoutEspace(vsIndexDeLivraison, 11) &
                     CHR_sAjoutEspace(vsNumLP, 3)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then
                    API_TID_bValideLP = True
                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

End Module


