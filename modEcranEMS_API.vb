Option Strict Off
Option Explicit On
Module modEcranEMS_API

    'Recherche de l'emplacement
    'API=MMS010MI
    'Fonction=GetLocation
    Public Function API_EMS_bRechercheEmplacement(ByVal vsDepot As String, ByVal vsEmplacement As String) As Boolean

        API_EMS_bRechercheEmplacement = False
        Dim sParam As String = ""
        Dim sResultat As String = ""

        If API_bConnexionAPI("MMS010MI") Then

            'Construction de la fonction et de ses paramètres pour l'appel API
            sParam = CHR_sAjoutEspace("GetLocation", 15) &
                     CHR_sAjoutEspace(gTab_Configuration.sSociete, 3) &
                     CHR_sAjoutEspace(vsDepot, 3) &
                     CHR_sAjoutEspace(vsEmplacement, 10)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then
                    API_EMS_bRechercheEmplacement = True
                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

    'Recherche Article/lot
    'API=MMS235MI
    'Fonction=GetItmLot
    Public Function API_EMS_bControleArticleLot(ByRef vsNumOrdre As String, ByVal vsArticle As String, ByVal vsLot As String) As Boolean

        API_EMS_bControleArticleLot = False
        Dim sParam As String = ""
        Dim sResultat As String = ""

        If API_bConnexionAPI("MMS235MI") Then

            'Construction de la fonction et de ses paramètres pour l'appel API
            sParam = CHR_sAjoutEspace("GetItmLot", 15) &
                     CHR_sAjoutEspace(vsArticle, 15) &
                     CHR_sAjoutEspace(vsLot, 20)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then

                    If (vsNumOrdre = "" Or (vsNumOrdre <> "" And vsNumOrdre = Trim(Mid(sResultat, 60, 10)))) Then
                        vsNumOrdre = Trim(Mid(sResultat, 60, 10))
                        API_EMS_bControleArticleLot = True
                    Else
                        MSG_AfficheErreur(giERR_CAB_NUM_ORDRE_INVALIDE)
                    End If

                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

    'Recherche de la quantité Article/lot
    'API=MMS060MI
    'Fonction=LstLot
    Public Function API_EMS_bControleQuantiteArticleLot(ByVal vsDepot As String, ByVal vsArticle As String, ByVal vsLot As String, ByVal vsQuantite As String, Optional ByVal vsEmplacement As String = "",
                                                        Optional ByRef vsQuantiteEnStock As String = "") As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""
        Dim nListe As Long
        Dim nQuantite As Long = 0
        Dim nQuantiteStock As Long = 0

        API_EMS_bControleQuantiteArticleLot = False
        vsQuantiteEnStock = ""

        If (IsNumeric(CHR_TransformeSeparateurPourNumerique(Trim(vsQuantite)))) Then
            nQuantite = (CHR_TransformeSeparateurPourNumerique(Trim(vsQuantite)))
        Else
            MSG_AfficheErreur(giERR_FORMAT_NUMERIC)
            Exit Function
        End If

        If API_bConnexionAPI("MMS060MI") Then

            'Construction de la fonction et de ses paramètres pour l'appel API
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

                        API_RetourneResultatSuivantPourREP(sResultat, nListe)
                    End While

                    If (nQuantite > nQuantiteStock) Then
                        MSG_AfficheErreur(giERR_QTE_INVALIDE)
                    Else
                        vsQuantiteEnStock = CStr(nQuantiteStock)
                        API_EMS_bControleQuantiteArticleLot = True
                    End If

                Else
                    MSG_AfficheErreur(giERR_ID_STOCK_INVALIDE)
                End If
            End If

        End If

    End Function

    'Edition de l'étiquette palette
    'API=MMS060MI
    'Fonction=PrtPutAwayLbl
    Public Function API_EMS_bEditionEtiquettePalette(ByVal vsDepot As String, ByVal vsEmplacement As String, ByVal vsNumOrdre As String, ByVal vsArticle As String, ByVal vsLot As String, ByVal vsQuantite As String) As Boolean

        API_EMS_bEditionEtiquettePalette = False
        Dim sParam As String = ""
        Dim sResultat As String = ""

        If API_bConnexionAPI("MMS060MI") Then

            'Construction de la fonction et de ses paramètres pour l'appel API
            sParam = CHR_sAjoutEspace("PrtPutAwayLbl", 15) &
                     CHR_sAjoutEspace(vsDepot, 3) &
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

                    API_EMS_bEditionEtiquettePalette = True

                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

    'Recherche de l'article
    'API=MMS200MI
    'Fonction=Get
    Public Function API_EMS_bRechercheArticle(ByVal vsArticle As String, ByRef vsArticleLibelle As String) As Boolean

        API_EMS_bRechercheArticle = False
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
                    API_EMS_bRechercheArticle = True
                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

    'Recherche du code motif
    'API=CRS175MI
    'Fonction=GetGeneralCode
    Public Function API_EMS_bControleCodeMotif(ByVal vsCodeMotif As String, ByRef vsCodeMotifLibelle As String) As Boolean


        Dim sParam As String = ""
        Dim sResultat As String = ""

        API_EMS_bControleCodeMotif = False
        vsCodeMotifLibelle = ""

        If API_bConnexionAPI("CRS175MI") Then

            'Construction de la fonction et de ses paramètres pour l'appel API
            sParam = CHR_sAjoutEspace("GetGeneralCode", 15) &
                     CHR_sAjoutEspace("", 3) &
                     CHR_sAjoutEspace("", 3) &
                     CHR_sAjoutEspace("RSCD", 10) &
                     CHR_sAjoutEspace(vsCodeMotif, 10)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then
                    vsCodeMotifLibelle = Trim(Mid(sResultat, 36, 15))
                    API_EMS_bControleCodeMotif = True
                Else
                    MSG_AfficheErreur(giERR_CODE_MOTIF_INVALIDE)
                End If
            End If

        End If

    End Function

    'Recherche de l'article
    'API=MMS310MI
    'Fonction=Update
    Public Function API_EMS_bAjustementDuStock(ByVal vsDepot As String, ByVal vsEmplacement As String, ByVal vsArticle As String, ByVal vsLot As String, ByVal vsQuantite As String, ByVal vsNouvelleQuantite As String,
                                               ByVal vsQuantiteEnStock As String, ByVal vsCodeMotif As String) As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""
        Dim nQuantiteEnStock As Long = 0
        Dim nQuantite As Long = 0
        Dim nNouvelleQuantite As Long = 0

        API_EMS_bAjustementDuStock = False

        'Calcul de la nouvelle quantité en stock sur Emplacement/Article/Lot
        If (IsNumeric(CHR_TransformeSeparateurPourNumerique(Trim(vsQuantite)))) Then
            nQuantite = CHR_TransformeSeparateurPourNumerique(Trim(vsQuantite))
        End If
        If (IsNumeric(CHR_TransformeSeparateurPourNumerique(Trim(vsNouvelleQuantite)))) Then
            nNouvelleQuantite = CHR_TransformeSeparateurPourNumerique(Trim(vsNouvelleQuantite))
        End If
        If (IsNumeric(CHR_TransformeSeparateurPourNumerique(Trim(vsQuantiteEnStock)))) Then
            nQuantiteEnStock = CHR_TransformeSeparateurPourNumerique(Trim(vsQuantiteEnStock))
        End If

        If (nNouvelleQuantite = nQuantite) Then
            API_EMS_bAjustementDuStock = True
            Exit Function
        Else
            nQuantiteEnStock = nQuantiteEnStock - (nQuantite - nNouvelleQuantite)
        End If



        If API_bConnexionAPI("MMS310MI") Then

            'Construction de la fonction et de ses paramètres pour l'appel API
            sParam = CHR_sAjoutEspace("Update", 15) &
                     CHR_sAjoutEspace(gTab_Configuration.sSociete, 3) &
                     CHR_sAjoutEspace(vsDepot, 3) &
                     CHR_sAjoutEspace(vsArticle, 15) &
                     CHR_sAjoutEspace(vsEmplacement, 10) &
                     CHR_sAjoutEspace(vsLot, 20) &
                     CHR_sAjoutEspace("", 30) &
                     CHR_sAjoutEspace(CStr(nQuantiteEnStock), 11) &
                     CHR_sAjoutEspace("", 106) &
                     CHR_sAjoutEspace(vsCodeMotif, 3)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then
                    API_EMS_bAjustementDuStock = True
                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function


End Module


