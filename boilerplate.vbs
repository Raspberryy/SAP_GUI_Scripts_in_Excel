Private Sub boilerPate()
    
' ********************
' **    GLOBALS     **
' ********************

    Dim sapProdDesc As String
    sapProdDesc = "INSERT_YOUR_PROD_DESCRIPTION_HERE"
    
    Dim SapGuiAuto As Object
    Dim SAP_APP As Object
    Dim connection As Object
 
' ****************************
' **    CONNECT TO SAP      **
' ****************************
    
    ' Check SAP is running
    On Error GoTo SAPNotRunning
        Set SapGuiAuto = GetObject("SAPGUI")
        Set SAP_APP = SapGuiAuto.GetScriptingEngine
    
SAPNotRunning:
    ' Start SAP.exe if needed
    If SapGuiAuto Is Nothing Then
        Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", vbNormalFocus
        Set WSHShell = CreateObject("WScript.Shell")
        Do Until WSHShell.AppActivate("SAP Logon ")
            Application.Wait Now + TimeValue("0:00:01")
            Loop
        Set WSHShell = Nothing
        
        Set SapGuiAuto = GetObject("SAPGUI")
        Set SAP_APP = SapGuiAuto.GetScriptingEngine
    End If
    
    ' Get session to PROD environment
    Dim i As Long
    
    If SAP_APP.Connections.Count() > 0 Then
        For i = 0 To SAP_APP.Connections.Count() - 1
            Set connection = SAP_APP.Children(CLng(i))
            If connection.Description() = sapProdDesc Then
                Exit For
            End If
        Next
    End If
    
    ' Start session to PROD environment if needed
    If connection Is Nothing Then
        Set connection = SAP_APP.OpenConnection(sapProdDesc)
    End If
    
' ***************************
' **    EXECUTE SCRIPT     **
' ***************************
    
    If Not IsObject(session) Then
        Set session = connection.Children(0)
    End If
    
    If IsObject(WScript) Then
        WScript.ConnectObject session, "on"
        WScript.ConnectObject Application, "on"
    End If
    
    ' Check if User is in any menu
    Set CancelButton = session.findById("wnd[0]/tbar[0]/btn[12]")
    If CancelButton.Changeable = True Then
        If MsgBox("Do you have any unsaved work in your SAP Client?", vbQuestion + vbYesNo + vbDefaultButton2, "Unsaved Work") = vbNo Then
            ' Close the stuff User was working on and return to Home Menu
			CancelButton.press
        Else
			' Script ends because User needs to save his stuff first
            MsgBox ("Please save changes and submit request again")
            Exit Sub
        End If
    End If
    
	' *******************************************************************************
	' *** INSERT GUI SCRIPT RECORDING BELOW *** INSERT GUI SCRIPT RECORDING BELOW ***
	' *******************************************************************************
	

    
	
	' *******************************************************************************
	' *** INSERT GUI SCRIPT RECORDING ABOVE *** INSERT GUI SCRIPT RECORDING ABOVE ***
	' *******************************************************************************   
    
	' Minimize SAP Client
    session.findById("wnd[0]").iconify
    
' *********************
' **    CLEAN UP     **
' *********************

    Set connection = Nothing
    
End Sub