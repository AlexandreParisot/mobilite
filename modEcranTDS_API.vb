Option Strict Off
Option Explicit On
Module modEcranTDS_API

    'Recherche de l'emplacement
    'API=MMS010MI
    'Fonction=GetLocation
    Public Function API_TDS_bRechercheEmplacement(ByVal vsEmplacement As String, Optional ByRef vsTypeEmplacement As String = "") As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""

        API_TDS_bRechercheEmplacement = False
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
                    API_TDS_bRechercheEmplacement = True
                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

    'Recherche Article/lot
    'API=MMS235MI
    'Fonction=GetItmLot
    Public Function API_TDS_bControleArticleLot(ByRef vsNumOrdre As String, ByVal vsArticle As String, ByVal vsLot As String) As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""
        API_TDS_bControleArticleLot = False

        If API_bConnexionAPI("MMS235MI") Then

            'Construction de la fonction et de ses paramètAPI pour l'appel API
            sParam = CHR_sAjoutEspace("GetItmLot", 15) &
                     CHR_sAjoutEspace(vsArticle, 15) &
                     CHR_sAjoutEspace(vsLot, 20)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then

                    If (vsNumOrdre = "" Or (vsNumOrdre <> "" And vsNumOrdre = Trim(Mid(sResultat, 60, 10)))) Then
                        vsNumOrdre = Trim(Mid(sResultat, 60, 10))
                        API_TDS_bControleArticleLot = True
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
    Public Function API_TDS_bControleQuantiteArticleLot(ByVal vsArticle As String, ByVal vsLot As String, ByVal vsQuantite As String, Optional ByVal vsEmplacement As String = "",
                                                        Optional ByRef vsStatutIDStock As String = "", Optional ByRef vsQuantiteAffectee As String = "", Optional ByRef vsDatePeremption As String = "") As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""
        Dim nListe As Long
        Dim nQuantite As Long = 0
        Dim nQuantiteStock As Long = 0

        API_TDS_bControleQuantiteArticleLot = False
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
                        API_TDS_bControleQuantiteArticleLot = True
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
    Public Function API_TDS_bRechercheArticle(ByVal vsArticle As String, ByRef vsArticleLibelle As String) As Boolean

        API_TDS_bRechercheArticle = False
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
                    API_TDS_bRechercheArticle = True
                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

    'Transfert du stock
    'API=MMS175MI
    'Fonction=Update
    Public Function API_TDS_bTransfertIdStock(ByVal vsEmplacementDeDebut As String, ByVal vsArticle As String, ByVal vsLot As String, ByVal vsQuantite As String,
                                            ByVal vsEmplacementDeFin As String) As Boolean

        Dim sParam As String = ""
        Dim sResultat As String = ""

        API_TDS_bTransfertIdStock = False

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
                    API_TDS_bTransfertIdStock = True
                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If

    End Function

End Module


