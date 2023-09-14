Option Strict Off
Option Explicit On
Module modConfiguration
	
	' Modules de test de la configuration
	' Test si les r�pertoires existent
	Public Sub CFG_VerifieArborescenceApplication()
        Dim sDir As String = ""

        'G�n�ral pour les logs
        sDir = My.Application.Info.DirectoryPath & gCST_sREPERTOIRE_LOG
        sDir = sDir.Substring(0, sDir.Length - 1)
        If Dir(sDir, FileAttribute.Directory) = "" Then
            MkDir(sDir)
        End If

    End Sub
End Module