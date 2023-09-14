Option Strict Off
Option Explicit On
Module modEcranEEP_API

    'Recherche de l'emplacement
    'API=MMS005MI
    'Fonction=GetWarehouse
    Public Function API_EEP_bRechercheDepot(ByVal vsDepot As String, ByRef vsDepotLibelle As String) As Boolean

        API_EEP_bRechercheDepot = False
        Dim sParam As String = ""
        Dim sResultat As String = ""
        vsDepotLibelle = ""

        If API_bConnexionAPI("MMS005MI") Then

            'Construction de la fonction et de ses paramètres pour l'appel API
            sParam = CHR_sAjoutEspace("GetWarehouse", 15) &
                     CHR_sAjoutEspace(vsDepot, 3)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then
                    vsDepotLibelle = Trim(Mid(sResultat, 28, 36))
                    API_EEP_bRechercheDepot = True
                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

    'Recherche de l'emplacement
    'API=MMS010MI
    'Fonction=GetLocation
    Public Function API_EEP_bRechercheEmplacement(ByVal vsDepot As String, ByVal vsEmplacement As String) As Boolean

        API_EEP_bRechercheEmplacement = False
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
                    API_EEP_bRechercheEmplacement = True
                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

    'Recherche Article/lot
    'API=MMS235MI
    'Fonction=GetItmLot
    Public Function API_EEP_bControleArticleLot(ByRef vsNumOrdre As String, ByVal vsArticle As String, ByVal vsLot As String) As Boolean

        API_EEP_bControleArticleLot = False
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
                        API_EEP_bControleArticleLot = True
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
    Public Function API_EEP_bControleQuantiteArticleLot(ByVal vsDepot As String, ByVal vsArticle As String, ByVal vsLot As String, ByVal vsQuantite As String, Optional ByVal vsEmplacement As String = "") As Boolean

        API_EEP_bControleQuantiteArticleLot = False
        Dim sParam As String = ""
        Dim sResultat As String = ""
        Dim nListe As Long
        Dim nQuantite As Long = 0
        Dim nQuantiteStock As Long = 0

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
                        API_EEP_bControleQuantiteArticleLot = True
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
    Public Function API_EEP_bEditionEtiquettePalette(ByVal vsDepot As String, ByVal vsEmplacement As String, ByVal vsNumOrdre As String, ByVal vsArticle As String, ByVal vsLot As String, ByVal vsQuantite As String) As Boolean

        API_EEP_bEditionEtiquettePalette = False
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

                    API_EEP_bEditionEtiquettePalette = True

                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

    'Recherche de l'article
    'API=MMS200MI
    'Fonction=Get
    Public Function API_EEP_bRechercheArticle(ByVal vsArticle As String, ByRef vsArticleLibelle As String) As Boolean

        API_EEP_bRechercheArticle = False
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
                    API_EEP_bRechercheArticle = True
                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function


End Module


