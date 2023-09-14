Option Strict Off
Option Explicit On
Module modEcranIntro
	
	
	Public Sub EcranIntro()
		Dim sScan As String
		
		With go_IO
			.RFPrint(0, 0, CHR_sCentrer(" MODULE ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)

            .RFPrint(0, 4, CHR_sCentrer("< RADIO >"), WirelessStudioOle.RFIOConstants.WLNORMAL)

            .RFPrint(0, 8, CHR_sCentrer("V" & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Revision), WirelessStudioOle.RFIOConstants.WLNORMAL)
			
			.RFPrint(0, go_TRM.TerminalHeight - 1, CHR_sCentrer("Appuyez sur ENTREE"), WirelessStudioOle.RFIOConstants.WLNORMAL)
			
			sScan = .GetEventEx(gCST_sFICHIER_BOUTONS_OK)
			
		End With
	End Sub
	
	Public Sub EcranTMPPourTEST(ByRef vsTexte As String)
		Dim sScan As String
		
		With go_IO
			.RFPrint(0, 0, CHR_sCentrer(" ECRAN DE TEST ", "="), WirelessStudioOle.RFIOConstants.WLREVERSE + WirelessStudioOle.RFIOConstants.WLCLEAR)

            .RFPrint(0, 3, CHR_sCentrer("< RADIO >"), WirelessStudioOle.RFIOConstants.WLNORMAL)

            .RFPrint(0, 6, CHR_sCentrer(vsTexte), WirelessStudioOle.RFIOConstants.WLNORMAL)
			
			.RFPrint(0, go_TRM.TerminalHeight - 1, CHR_sCentrer("Appuyez sur ENTREE"), WirelessStudioOle.RFIOConstants.WLNORMAL)
			
			sScan = .GetEventEx(gCST_sFICHIER_BOUTONS_OK)
			
		End With
	End Sub
End Module