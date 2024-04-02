Option Strict Off
Option Explicit On
Module modMain


    'Entrée principale du programme
    Public Sub Main()

        SEC_InitProtectionMenu()

        CFG_VerifieArborescenceApplication()

        InitVariableGlobale()
        Try
            'Recupération de la configuration du terminal
            'Nbre de ligne, nbre de colonne ...
            If bWRL_RecupereConfiguration() Then

                'Récupération des différents paramètrages
                If bINI_InitParametrage() Then
                    'Ecriture du log d'initialisation du PDT
                    If bLDF_LogInitialisationPDT() Then
                        'Initialisation du code barre du PDT
                        If bWRL_InitCodeBarrePDT() Then

                            EcranIntro()

                            'Saisie de l'utilisateur
                            If bELG_SaisieLogin() Then
                                'Initialisation des menus

                                If bMNU_InitMenu() Then

                                    EcranPrincipal()
                                Else
                                    LDF_LogErreurApplication(giPROC_bInitMenu, False)
                                End If
                            End If
                            FinApplication()
                        Else
                            LDF_LogErreurApplication(giPROC_bInitCodeBarrePDT, False)
                        End If
                    End If
                Else
                    LDF_LogErreurApplication(giPROC_bInitParametrage, False)
                End If
            Else
                LDF_LogErreurApplication(giPROC_bRecupereConfiguration, False)
            End If

        Catch ex As Exception
            LDF_AfficheErreurDansLog("1", "0", ex.ToString)
        End Try


    End Sub

    Private Sub FinApplication()
        go_BAR.DeleteBarcodeFile(gCST_sFICHIER_CODE_BARRE)
    End Sub


    'Au début de l'application, certaines variables ont besoin d'être initialisées par défaut
    Private Sub InitVariableGlobale()

        gTab_General.sFichierErr = My.Application.Info.DirectoryPath & gCST_sREPERTOIRE_LOG & "ErreurCOM.log"
        gTab_Configuration.bLog = True

    End Sub
End Module