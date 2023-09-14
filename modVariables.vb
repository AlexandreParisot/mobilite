Option Strict Off
Option Explicit On
Module modVariables

    Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer


    '*******************************************************************************
    '* VARIABLES GENERALES
    '*******************************************************************************
    'Repertoire des Logs
    Public Const gCST_sREPERTOIRE_LOG As String = "\LOG\"

    'Constantes
    Public Const gCST_sFICHIER_MNU_GENERAL As String = "MnuGene"
    Public Const gCST_sFICHIER_MNU_SORTIE As String = "MnuSSto"
    Public Const gCST_sFICHIER_CODE_BARRE As String = "CodeBar"
    Public Const gCST_sFICHIER_BOUTONS_OK_CLR_QUIT As String = "Btn001"
    Public Const gCST_sFICHIER_BOUTONS_OK_CLR As String = "Btn002"
    Public Const gCST_sFICHIER_BOUTONS_OK As String = "Btn003"
    Public Const gCST_sFICHIER_BOUTONS_OK_CLR_QUIT_SAISIE As String = "Btn004"
    Public Const gCST_sFICHIER_BOUTONS_OK_CLR_PRECEDENT_SUIVANT As String = "Btn005"
    Public Const gCST_sFICHIER_BOUTONS_FPAL_FLIG_CLR As String = "Btn006"
    Public Const gCST_sFICHIER_BOUTONS_VALIDATION_ANNULATION As String = "Btn007"
    Public Const gCST_sFICHIER_BOUTONS_OK_CLR_LIG_SUIVANTE_M_ATT As String = "Btn008"
    Public Const gCST_sFICHIER_BOUTONS_FIN_ABSENT_AJOUT As String = "Btn009"
    Public Const gCST_sFICHIER_BOUTONS_OK_CLR_RETOUR As String = "Btn010"
    Public Const gCST_sFICHIER_BOUTONS_FIN_AJOUT As String = "Btn011"
    Public Const gCST_sFICHIER_BOUTONS_OK_CLR_LIG_SUIVANTE_M_ATT_FPAL As String = "Btn012"
    Public Const gCST_sFICHIER_BOUTONS_FPAL_FLIG_CLR_CHGLOT As String = "Btn013"
    Public Const gCST_sFICHIER_BOUTONS_OK_CLR_QUIT_PRECEDENT_SUIVANT_SAISIE_FIN_LIGNE As String = "Btn014"
    Public Const gCST_sFICHIER_BOUTONS_OK_CLR_QUIT_F_PAL As String = "Btn015"
    Public Const gCST_sFICHIER_BOUTONS_OK_QUIT_PRECEDENT_SUIVANT As String = "Btn016"
    Public Const gCST_sFICHIER_BOUTONS_OK_CLR_QUIT_SUIVANT As String = "Btn0017"
    Public Const gCST_sFICHIER_BOUTONS_OK_CLR_QUIT_IMPRIME_DERNIERE_ETIQUETTE As String = "Btn0018"

    'Variable globale à l'application déterminant si il y a eu une erreur de communication
    'Avec un PDT auquel cas l'application doit  etre fermée !
    Public gbErreurCommunication As Boolean

    Public gTab_Configuration As Enr_Configuration
    Public gTab_General As Enr_General
    Public gTab_Menu As Enr_Menu
    Public gtab_Securite As Enr_Securite

    'Gestion des messages
    Public gTab_Messages(12) As String


    '*******************************************************************************
    '* VARIABLES DESTINEES AU MODULE modRPF_FicIni
    '*******************************************************************************
    Public Const gCST_sFICHIER_INI As String = "GstSto.ini"
    Public Const gCST_sFICHIER_MENU_INI As String = "Menu.ini"

    'Constantes des clefs pour GstSto.ini
    Public Const gCST_INI_SEC_API As String = "API"
    Public Const gCST_INI_SEC_PDT As String = "PDT"
    Public Const gCST_INI_SEC_M3 As String = "M3"
    Public Const gCST_INI_SEC_APP As String = "APP"
    Public Const gCST_INI_SEC_EEP As String = "EEP"
    Public Const gCST_INI_SEC_ABC As String = "ABC"
    Public Const gCST_INI_SEC_RES As String = "RES"
    Public Const gCST_INI_SEC_TDS As String = "TDS"
    Public Const gCST_INI_SEC_EMS As String = "EMS"
    Public Const gCST_INI_SEC_TID As String = "TID"
    Public Const gCST_INI_SEC_USER As String = "USER"

    'API
    Public Const gCST_INI_IP As String = "IP"
    Public Const gCST_INI_PORT As String = "PORT"
    Public Const gCST_INI_DOMAINE As String = "DOMAINE"

    'PDT
    Public Const gCST_INI_DLA As String = "DLA"
    Public Const gCST_INI_CODBAR As String = "CODBAR"

    'M3
    Public Const gCST_INI_CONO As String = "CONO"
    Public Const gCST_INI_DIVI As String = "DIVI"
    Public Const gCST_INI_WHLO As String = "WHLO"
    Public Const gCST_INI_TYPE_EMPLACEMENT_BORD_DE_CHAINE_NORMAL As String = "TYPE_EMPLACEMENT_BORD_DE_CHAINE_NORMAL"
    Public Const gCST_INI_TYPE_EMPLACEMENT_BORD_DE_CHAINE_BIB As String = "TYPE_EMPLACEMENT_BORD_DE_CHAINE_BIB"

    'RES
    Public Const gCST_INI_RES_EMPLACEMENT_FINAL_POUR_TB As String = "EMPLACEMENT_FINAL_POUR_TB"

    'TID


    'APP
    Public Const gCST_INI_TIMW As String = "TIMW"
    Public Const gCST_INI_LOG As String = "LOG"

    'Constantes des clefs pour Menu.ini
    Public Const gCST_INI_SEC_MENU As String = "MENU"

    'Structure des utilisateurs paramètrés dans GstSto.ini
    Public Structure Enr_Users
        Dim sUtilisateur As String
        Dim sMotDePasse As String
    End Structure

    'Structure du fichier GstSto.Ini
    Public Structure Enr_Configuration
        Dim sIP As String
        Dim sPort As String
        Dim sDomaine As String
        Dim sUtilisateur As String
        Dim sMotDePasse As String
        Dim iDelaiMessage As Short
        Dim sCodeBar As String
        Dim sSociete As String
        Dim sDivision As String
        Dim sDepot As String
        Dim lTimeWait As Integer
        Dim bLog As Boolean
        Dim sSLPT_Bord_De_Chaine_Normal As String
        Dim sSLPT_Bord_De_Chaine_Bib As String
        Dim sRES_WHSL_TB_Final As String
        Dim sProfil() As Enr_Users
    End Structure

    Public Structure Enr_General
        Dim sFichierLog As String
        Dim sFichierErr As String
        Dim sPDT As String
        Dim sIP_PDT As String
        Dim sFichierAPI As String
    End Structure

    Public Structure Enr_Menu
        Dim nNombreMenu As Short
        Dim Tab_Menu() As String 'Nom des menus
        Dim Tab_DroitMenu() As Boolean 'Droit Sur Menu
        Dim Tab_Option() As String
    End Structure

    'Tableaux de gestion Sécurité
    Public Structure Enr_Securite
        Dim bEEP As Boolean
        Dim bABC As Boolean
        Dim bRES As Boolean
        Dim bTDS As Boolean
        Dim bEMS As Boolean
        Dim bTID As Boolean
        Dim bPI1 As Boolean
    End Structure
End Module