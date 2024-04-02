Option Strict Off
Option Explicit On
Module modAPI_M3

    'PREFIXE affecté au module = "API"

    'Module de traitement et d'appel des API M3
    Dim Sock As New MVXSOCKX_SVRLib.MvxSockX

    Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Integer)

    '********************************************************************************
    '* LES DEUX FONCTIONS PRINCIPALES D'APPEL DES APIs
    '********************************************************************************

    'Fonction générale qui retourne True si la Connexion à l'API
    's'est correctement déroulée
    'Paramètre:
    'Le code de l'API
    Public Function API_bConnexionAPI(ByRef vsCodeAPI As String) As Boolean
        Dim sAPI As String = ""
        Dim sErreur As String = ""
        Dim rc As Integer = 0

        Try
            'Affichage sur PDT
            LDF_AfficheErreurDansLog("1", "0", "2")
            WRL_AfficheTraitementEnCours(vsCodeAPI)
            LDF_AfficheErreurDansLog("1", "0", "3")
            With gTab_Configuration
                sAPI = sAPI & vsCodeAPI

                'Fermeture du socket
                LDF_LogPourTrace("Fermeture socket API...")

                Sock.MvxSockClose()

                Sleep(.lTimeWait)

                LDF_LogPourTrace("Connexion socket à " & vsCodeAPI)

                rc = Sock.MvxSockConnect(.sIP, CShort(.sPort), .sDomaine + "\" + .sUtilisateur, .sMotDePasse, sAPI, "")

                LDF_LogPourTrace("Etat connexion à " & vsCodeAPI & " : " & rc)

                'Attente
                Sleep(.lTimeWait)

                If rc <> 0 Then
                    Sock.MvxSockGetLastError(sErreur)
                    MSG_AfficheErreur(giERR_INIT_PROCEDURE_API, CHR_sSupAccent(sErreur))
                    Sock.MvxSockClose()
                Else
                    API_bConnexionAPI = True
                End If

            End With

        Catch ex As Exception
            LDF_AfficheErreurDansLog("1", "0", ex.ToString)
        End Try


    End Function

    'Fonction de traitement de la fonction API et retourne True si le traitement a abouti
    'Et renvoie le résultat
    'Paramètres:
    '-vsParam as string
    '-vsResultat as string
    '-vbNotClose as boolean : Si la fonction est appelée pour récupérer une liste ou rester connecté à la même API,
    'on ne ferme pas la connexion pour la travailler...
    Public Function API_bTraitementAPI(ByRef vsParam As String, ByRef vsResultat As String, Optional ByRef vbNotClose As Boolean = False) As Boolean
        Dim rc As Integer
        Dim sErreur As String = ""

        LDF_LogPourTrace("Appel API avec : " & vsParam)

        rc = Sock.MvxSockTrans(vsParam, vsResultat)

        'Attente
        Sleep(gTab_Configuration.lTimeWait)

        LDF_LogPourTrace("Retour : " & rc & "=>" & vsResultat)

        vsResultat = CHR_sSupAccent(vsResultat)

        If rc <> 0 Then
            Sock.MvxSockGetLastError(sErreur)
            MSG_AfficheErreur(giERR_PROCEDURE_API, CHR_sSupAccent(sErreur), vsResultat)
        Else
            API_bTraitementAPI = True
        End If

        If Not vbNotClose Then
            Sock.MvxSockClose()
        End If

        'Attente
        Sleep(gTab_Configuration.lTimeWait)

    End Function

    Public Sub API_FermeConnexionPourREP()
        Sock.MvxSockClose()
    End Sub

    Public Sub API_RetourneResultatSuivantPourREP(ByRef vsResultat As String, ByRef vnIndex As Short)
        Dim rc As Integer
        Dim sErreur As String = ""

        vnIndex = vnIndex + 1

        'Affichage sur PDT
        WRL_AfficheTraitementEnCours("Liste " & vnIndex & "...")

        rc = Sock.MvxSockReceive(vsResultat)

        LDF_LogPourTrace("Retour : " & rc & "=>" & vsResultat)

        vsResultat = CHR_sSupAccent(vsResultat)

        If rc <> 0 Then
            Sock.MvxSockGetLastError(sErreur)
            MSG_AfficheErreur(giERR_PROCEDURE_API, CHR_sSupAccent(sErreur), vsResultat)
            Sock.MvxSockClose()
        End If

    End Sub

    'Procédure Spécifique
    'API=GENERAL
    'Proc=GetServerTime
    Public Function API_bRecupDateHeure(ByRef vsDateCourante As String, ByRef vsHeureCourante As String) As Boolean

        Dim sParam As String
        Dim sResultat As String = ""
        API_bRecupDateHeure = False

        If API_bConnexionAPI("GENERAL") Then

            'Construction de la fonction et de ses paramètres pour l'appel API
            sParam = CHR_sAjoutEspace("GetServerTime", 15)

            If API_bTraitementAPI(sParam, sResultat) Then
                If Mid(sResultat, 1, 3) <> "NOK" Then
                    vsDateCourante = Trim(Mid(sResultat, 16, 10))
                    vsHeureCourante = Trim(Mid(sResultat, 26, 6))
                    API_bRecupDateHeure = True
                Else
                    MSG_AfficheErreur(giERR_PROCEDURE_API, sResultat)
                End If
            End If

        End If
    End Function

End Module