Option Strict Off
Option Explicit On
Module modCaracteres
	
	
	'PREFIXE affect� au module = "CHR"
	
	
	'Fonction qui renvoie la chaine de caract�res en param�tre centr�e
	'Par rapport au nombre de caract�res du PDT
	'ATTENTION : Ne convient pas pour les menus
	Public Function CHR_sCentrer(ByRef vsChaine As String, Optional ByRef vsAjout As String = " ") As String

        CHR_sCentrer = ""
        Dim iTaille As Short
		Dim iNbrSpc As Short
		
		iTaille = Len(vsChaine)
		If iTaille > 0 Then
			iNbrSpc = (go_TRM.TerminalWidth - iTaille) \ 2
			If iNbrSpc > 0 Then
				CHR_sCentrer = New String(vsAjout, iNbrSpc) & vsChaine & New String(vsAjout, iNbrSpc)
				If ((iNbrSpc * 2) + iTaille) < go_TRM.TerminalWidth Then
					CHR_sCentrer = CHR_sCentrer & vsAjout
				End If
			End If
		End If
		
	End Function
	
	'Fonction qui renvoie la chaine de caract�res en param�tre cadrer � droite
	'Par rapport au nombre de caract�res du PDT
	'ATTENTION : Ne convient pas pour les menus
	Public Function CHR_sCentrerDroite(ByRef vsChaine As String, Optional ByRef vsAjout As String = " ") As String
		
		Dim iTaille As Short
        Dim iNbrSpc As Short
        CHR_sCentrerDroite = ""

        iTaille = Len(vsChaine)
		If iTaille > 0 Then
			iNbrSpc = go_TRM.TerminalWidth - iTaille
			If iNbrSpc > 0 Then
				CHR_sCentrerDroite = New String(vsAjout, iNbrSpc) & vsChaine
			End If
		End If
		
	End Function

    'Fonction qui renvoie le N� de colonne de la chaine de caract�res en param�tre centr�e
    'Par rapport au nombre de caract�res du PDT
    'ATTENTION : Ne convient pas pour les menus
    Public Function CHR_nCentrer(ByRef vnLongueurChaine As Short) As Short
		
		If vnLongueurChaine > 0 Then
			CHR_nCentrer = (go_TRM.TerminalWidth - vnLongueurChaine) \ 2
		Else
			CHR_nCentrer = 0
		End If

    End Function

    'Fonction qui remplace les lettres avec accents pour les rendre
    'compatible avec DOS
    Public Function CHR_sSupAccent(ByRef vsChaine As String) As String
		Dim sChaine As String
		
		sChaine = Replace(vsChaine, "�", "e")
		
		sChaine = Replace(sChaine, "{", "e")
		
		sChaine = Replace(sChaine, "}", "e")
		
		sChaine = Replace(sChaine, "@", "a")
		
		sChaine = Replace(sChaine, "�", "e")
		
		sChaine = Replace(sChaine, "�", "o")
		
		sChaine = Replace(sChaine, "�", "a")
		
		sChaine = Replace(sChaine, "�", "i")
		
		CHR_sSupAccent = sChaine
		
	End Function

    'Retourne une chaine compl�t�e avec des espaces pour atteindre la taille voulue
    'Param�tres:
    '-Chaine de base
    'Taille en sortie
    'Option Espace en zero pour les valeurs numeriques
    Public Function CHR_sAjoutEspace(ByRef vsParametre As String, ByRef vnTaille As Short, Optional ByRef vbEspaceEnZero As Boolean = False, Optional ByRef vbEspaceDevant As Boolean = False) As String
		Dim iLng As Short
		
		iLng = Len(vsParametre)
		
		iLng = vnTaille - iLng
		
		If iLng >= 0 Then
			If vbEspaceEnZero Or vbEspaceDevant Then
				CHR_sAjoutEspace = Space(iLng) & vsParametre
			Else
				CHR_sAjoutEspace = vsParametre & Space(iLng)
			End If
		Else
			CHR_sAjoutEspace = Mid(vsParametre, 1, vnTaille)
		End If
		If vbEspaceEnZero Then
			CHR_sAjoutEspace = Replace(CHR_sAjoutEspace, " ", "0")
		End If
		
	End Function

    'Transforme le point �ventuel de la chaine en param�tre par une virgule
    'Pour la rendre compatible en numerique
    'Si le parametre vnNbrDecimal est pr�cis�, on v�rifie que la chaine n'en comporte pas plus
    Public Function CHR_TransformePointPourNumerique(ByRef vsCode As String, Optional ByRef vnNbrDecimal As Short = -1, Optional ByRef vbDemandeRetourErreur As Boolean = False) As String
        Dim sChaine As String
        Dim sChaine2 As String
        Dim nPos As Short
        Dim nIndex As Short
        Dim nDec As Short

        vbDemandeRetourErreur = False

        sChaine = Replace(vsCode, ".", ",")
        sChaine2 = sChaine
        If vnNbrDecimal > -1 Then
            nPos = InStr(1, sChaine, ",")
            If nPos > 0 Then
                sChaine2 = Mid(sChaine, 1, nPos)
                nDec = 1
                For nIndex = (nPos + 1) To Len(sChaine)
                    If nDec > vnNbrDecimal Then
                        vbDemandeRetourErreur = True
                        Exit For
                    End If
                    sChaine2 = sChaine2 & Mid(sChaine, nIndex, 1)
                    nDec = nDec + 1
                Next
            End If
        End If
        CHR_TransformePointPourNumerique = sChaine2

    End Function

    'Transforme le s�parateur d�cimal �ventuel de la chaine en s�parateur syst�me
    'Pour la rendre compatible en numerique
    'Si le parametre vnNbrDecimal est pr�cis�, on v�rifie que la chaine n'en comporte pas plus
    Public Function CHR_TransformeSeparateurPourNumerique(ByRef vsCode As String, Optional ByRef vnNbrDecimal As Short = -1, Optional ByRef vbDemandeRetourErreur As Boolean = False) As String
        Dim sChaine As String
        Dim sChaine2 As String
        Dim nPos As Short
        Dim nIndex As Short
        Dim nDec As Short
        Dim sSeparateurSystem As String
        Dim sSeparateurRecu As String
        Dim position_separateur As Integer

        vbDemandeRetourErreur = False

        sSeparateurSystem = System.Globalization.CultureInfo.InstalledUICulture.NumberFormat.NumberDecimalSeparator
        sSeparateurRecu = ""

        'Tester le '.'
        position_separateur = vsCode.IndexOf(".")
        If (position_separateur > 0) Then
            sSeparateurRecu = "."
        End If

        'Tester la ','
        position_separateur = vsCode.IndexOf(",")
        If (position_separateur > 0) Then
            sSeparateurRecu = ","
        End If

        sChaine = Replace(vsCode, sSeparateurRecu, sSeparateurSystem)
        sChaine2 = sChaine
        If vnNbrDecimal > -1 Then
            nPos = InStr(1, sChaine, sSeparateurSystem)
            If nPos > 0 Then
                sChaine2 = Mid(sChaine, 1, nPos)
                nDec = 1
                For nIndex = (nPos + 1) To Len(sChaine)
                    If nDec > vnNbrDecimal Then
                        vbDemandeRetourErreur = True
                        Exit For
                    End If
                    sChaine2 = sChaine2 & Mid(sChaine, nIndex, 1)
                    nDec = nDec + 1
                Next
            End If
        End If
        CHR_TransformeSeparateurPourNumerique = sChaine2

    End Function

    'Transforme la virgule en point pour les APIs
    Public Function CHR_TransformeVirguleEnPoint(ByRef vsCode As String) As String
		
		CHR_TransformeVirguleEnPoint = Replace(vsCode, ",", ".")
		
	End Function
	
	
	'Retourne un libell� en fonction de la taille de l'�cran et du niveau de ligne
	'Jusqu'� 4 Ligne
	Public Function CHR_sVerifieTaille(ByRef vsLibelle As String, ByRef vnNumLigne As Short) As String
		Dim sChaine As New VB6.FixedLengthString(120)
		
		sChaine.Value = vsLibelle
		
		CHR_sVerifieTaille = Mid(sChaine.Value, 1 + ((vnNumLigne - 1) * go_TRM.TerminalWidth), go_TRM.TerminalWidth)
		
	End Function

    'Retourne une chaine fabriqu�e du nombre de d�cimales pass�es en param�tre
    Public Function CHR_sRetourneQuantiteFormatDecimal(ByRef vnNbrDecimal As Short) As String
		Dim nIndex As Short
		Dim sChaine As String
		
		If vnNbrDecimal > 0 Then
			sChaine = "0,"
			For nIndex = 1 To vnNbrDecimal
				sChaine = sChaine & "0"
			Next 
		Else
			sChaine = "0"
		End If
		CHR_sRetourneQuantiteFormatDecimal = sChaine
		
	End Function

    'Fonction qui transforme une valeur en Double
    'en une valeur en alpha en supprimant la virgule
    Public Function CHR_sTransFormeDoubleEnAlpha(ByRef viNombre As Double) As String
		
		Dim sNombre As String
		Dim iPos As Short
		Dim sEntier As String
		Dim sDecimal As String
		
		sNombre = VB6.Format(viNombre, "0.000000")
		
		iPos = InStr(1, sNombre, ",")
		If iPos > 0 Then
			sEntier = Mid(sNombre, 1, iPos - 1)
			sDecimal = Mid(sNombre, iPos + 1, Len(sNombre) - iPos)
			CHR_sTransFormeDoubleEnAlpha = sEntier & sDecimal
		Else
			CHR_sTransFormeDoubleEnAlpha = sNombre
		End If
		
	End Function

    'Fonction qui transforme une valeur en Double
    'en une valeur en alpha en supprimant la virgule
    Public Function CHR_sTransFormeSingleEnAlpha(ByRef viNombre As Single) As String

        Dim sNombre As String
        Dim iPos As Short
        Dim sEntier As String
        Dim sDecimal As String

        sNombre = VB6.Format(viNombre, "0.000000")

        iPos = InStr(1, sNombre, ",")
        If iPos > 0 Then
            sEntier = Mid(sNombre, 1, iPos - 1)
            sDecimal = Mid(sNombre, iPos + 1, Len(sNombre) - iPos)
            CHR_sTransFormeSingleEnAlpha = sEntier & sDecimal
        Else
            CHR_sTransFormeSingleEnAlpha = sNombre
        End If

    End Function

    'Recherche des infos port�es par l'EAN128 saisi
    'Il se d�compose de la mani�re suivante, avec un identifiant entre parenth�se pour s�parer chaque valeur : 
    'N� Ordre (optionnel) = Indentifiant 21 + N�Ordre
    'Article = Identifiant 92 + Code Article
    'Lot = Identifiant 10 + Code Lot
    'Quantit� sur la palette = Identifiant 37 + Quantit�
    'Exemple : (21)234567(92)BOU000001(10)3L26518(37)36000 cela donne:
    ' N� ordre 234567, Article BOU000001, Lot 3L26518, Quantit� 36000
    Public Function CHR_bRecupInfoEAN128(ByVal vsEAN128Saisi As String, ByRef vsNumOrdre As String, ByRef vsArticle As String, ByRef vsLot As String, ByRef vsQuantite As String) As Boolean

        CHR_bRecupInfoEAN128 = False
        Dim nPosParentheseOuvrante As Short = 999
        Dim nPosParentheseFermante As Short = 999
        Dim nPosParentheseOuvrante_suivante As Short = 999
        Dim sValeur As String = ""
        Dim sIdentifiant As String = ""
        Dim sParentheseOuvrante = "("
        Dim sParentheseFermante = ")"

        vsNumOrdre = ""
        vsArticle = ""
        vsLot = ""
        vsQuantite = ""

        LDF_LogPourTrace("CAB EAN128 LU ===== : " & vsEAN128Saisi)


        'Recherche 1�re parenth�se ouvrante
        nPosParentheseOuvrante = InStr(1, vsEAN128Saisi, sParentheseOuvrante)

        'Recherche des valeurs contenues dans le CAB entre chaque identifiant
        While (nPosParentheseOuvrante > 0)

            'Recherche de la parenth�se fermante correspondante
            nPosParentheseFermante = InStr(nPosParentheseOuvrante + 1, vsEAN128Saisi, sParentheseFermante)

            'Recherche de la parenth�se ouvrante suivante
            nPosParentheseOuvrante_suivante = InStr(nPosParentheseFermante + 1, vsEAN128Saisi, sParentheseOuvrante)


            'R�cup�ration de la valeur
            If (nPosParentheseOuvrante_suivante > 0) Then
                sIdentifiant = Mid(vsEAN128Saisi, nPosParentheseOuvrante + 1, (nPosParentheseFermante - nPosParentheseOuvrante - 1))
                sValeur = Mid(vsEAN128Saisi, nPosParentheseFermante + 1, nPosParentheseOuvrante_suivante - (nPosParentheseFermante + 1))
                nPosParentheseOuvrante = nPosParentheseOuvrante_suivante
            Else
                sIdentifiant = Mid(vsEAN128Saisi, nPosParentheseOuvrante + 1, (nPosParentheseFermante - nPosParentheseOuvrante - 1))
                sValeur = Mid(vsEAN128Saisi, nPosParentheseFermante + 1, (vsEAN128Saisi.Length + 1) - (nPosParentheseFermante + 1))
                nPosParentheseOuvrante = 0
            End If

            If (sIdentifiant = 21) Then
                vsNumOrdre = sValeur
            End If
            If (sIdentifiant = 92) Then
                vsArticle = sValeur
            End If
            If (sIdentifiant = 10) Then
                vsLot = sValeur
            End If
            If (sIdentifiant = 37) Then
                vsQuantite = sValeur
            End If

        End While

        If (vsArticle <> "" And vsLot <> "" And vsQuantite <> "") Then
            CHR_bRecupInfoEAN128 = True
        Else
            MSG_AfficheErreur(giERR_CAB_INVALIDE)
        End If


    End Function

End Module