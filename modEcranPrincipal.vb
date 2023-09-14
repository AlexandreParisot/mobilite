Option Strict Off
Option Explicit On
Module modEcranPrincipal

    'Module avec définition de l'écran principal
    Public Sub EcranPrincipal()
        Dim nOptmnu As Short
        Dim bFin As Boolean

        While Not bFin And Not gbErreurCommunication

            'Remise à blanc de l'écran
            go_IO.RFPrint(0, 0, "", WirelessStudioOle.RFIOConstants.WLCLEAR)

            nOptmnu = go_MNU.DoMenu(gCST_sFICHIER_MNU_GENERAL)
            Select Case nOptmnu
                ' RF Error
                Case -2, -1
                    bFin = True

                Case Is < giOPT_FIN
                    EcranSelonMenu(nOptmnu)

                Case giOPT_FIN
                    bFin = True

                Case Else
                    MSG_AfficheErreur(giERR_OPTION_NON_PREVUE)

            End Select
        End While

    End Sub


    Private Sub EcranSelonMenu(ByRef vnOptMenu As Short)

        With gTab_Menu

            Select Case .Tab_Option(vnOptMenu)

                Case "EEP"
                    Ecran_EEP_EditionEtiquettePalette()
                Case "ABC"
                    Ecran_ABC_ApprovisionnementBordDeChaine()
                Case "RES"
                    Ecran_RES_RetourEnStock()
                Case "TDS"
                    Ecran_TDS_TransfertDeStock()
                Case "EMS"
                    Ecran_EMS_EtiquetteMAJStock()
                Case "TID"
                    Ecran_TID_TransfertInterDepot()
                Case "PI1"
                    Ecran_PI1_InventairePID()
            End Select

        End With
    End Sub

    Private Sub Ecran_EEP_EditionEtiquettePalette()
        EcranEEP("ETIQUETTE PALETTE")
    End Sub

    Private Sub Ecran_ABC_ApprovisionnementBordDeChaine()
        EcranABC("APPRO BORD DE CHAINE")
    End Sub

    Private Sub Ecran_RES_RetourEnStock()
        EcranRES("RETOUR EN STOCK")
    End Sub

    Private Sub Ecran_TDS_TransfertDeStock()
        EcranTDS("TRANSFERT DE STOCK")
    End Sub

    Private Sub Ecran_EMS_EtiquetteMAJStock()
        EcranEMS("ETIQUETTE MAJ STOCK")
    End Sub

    Private Sub Ecran_TID_TransfertInterDepot()
        EcranTID("TRANSF. INTER-DEPOT")
    End Sub
    Private Sub Ecran_PI1_InventairePID()
        EcranPI1("INVENTAIRE PID")
    End Sub

    'Affichage d'un écran non disponible pour l'option demandée
    Public Sub EcranNonDisponible(ByRef vsTitre As String)
        Dim sScan As String
        With go_IO
            .RFPrint(0, 0, CHR_sCentrer(" " & vsTitre & " ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)
            .RFPrint(0, 3, CHR_sCentrer("Non disponible"), WirelessStudioOle.RFIOConstants.WLNORMAL)

            sScan = .GetEvent()

        End With
    End Sub
End Module
