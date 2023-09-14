Option Strict Off
Option Explicit On
Module modSecurite

    'Les menus restent accessibles et param�trables. Cependant, on peut les interdire par le programme
    Public Sub SEC_InitProtectionMenu()
        With gtab_Securite
            .bEEP = True
            .bABC = True
            .bRES = True
            .bTDS = True
            .bEMS = True
            .bTID = True
            .bPI1 = True
        End With
    End Sub

    'Analyse si ce menu est autoris� pour ce client
    'On analyse les trois premiers caract�res du menus ( le menu provient du fichier Menu.ini Section :[MENU] )
    Public Function bSEC_AnalyseSecurite(ByRef vsMenu As String) As Boolean
		Dim sOption As String
		
		sOption = Mid(vsMenu, 1, 3)
		
		With gtab_Securite

            Select Case sOption
                Case "EEP"
                    If .bEEP Then
                        bSEC_AnalyseSecurite = True
                    End If
                Case "ABC"
                    If .bABC Then
                        bSEC_AnalyseSecurite = True
                    End If
                Case "RES"
                    If .bRES Then
                        bSEC_AnalyseSecurite = True
                    End If
                Case "TDS"
                    If .bTDS Then
                        bSEC_AnalyseSecurite = True
                    End If
                Case "EMS"
                    If .bEMS Then
                        bSEC_AnalyseSecurite = True
                    End If
                Case "TID"
                    If .bTID Then
                        bSEC_AnalyseSecurite = True
                    End If
                Case "PI1"
                    If .bPI1 Then
                        bSEC_AnalyseSecurite = True
                    End If
            End Select

        End With
		
	End Function
End Module