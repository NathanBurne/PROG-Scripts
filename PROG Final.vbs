Option Explicit

Call Start

Sub Start
	If isCScript = True Then
		Dim wshNet, wshShell, wshEnv, input
		Set wshNet = CreateObject("WScript.Network")
		Set wshShell = CreateObject("WScript.Shell")
		Set wshEnv = wshShell.Environment("System")
		
		WScript.StdOut.WriteLine "Hello " & wshNet.UserName & ", welcome to Scripting and Automation."
		WScript.StdOut.WriteLine "You have connected to " & wshNet.UserDomain & "\" & wshNet.ComputerName & " which is running " & wshEnv.Item("OS") & "."
		WScript.StdOut.WriteLine "You are running your script " & WScript.ScriptName & " from " & wshShell.CurrentDirectory & "."
		WScript.StdOut.WriteLine "The team members are: Jake Flemming(100521730)."
		WScript.StdOut.WriteLine
		
		WScript.StdOut.Write "Press ENTER to continue..."
	 	input = WScript.StdIn.ReadLine
		WScript.Timeout = 1800
		Call Menu
	Else
		MsgBox "This Project Must Be Run In CSCRIPT Only."
		WScript.Quit
	End If
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub Menu
	Dim input
	Do
		formatScreen(15)
		
		WScript.StdOut.WriteLine "		    Scripting and Automation - PROG 4103"
		WScript.StdOut.WriteLine "		   --------------------------------------"
		WScript.StdOut.WriteLine "		         A-Desktop Management"
		WScript.StdOut.WriteLine "		         B-Logon and Logoff"
		WScript.StdOut.WriteLine "		         C-User Management"
		WScript.StdOut.WriteLine "		         D-Disk Management"
		WScript.StdOut.WriteLine "		         E-Utilities"
		WScript.StdOut.WriteLine "		         Q-Quit Program"
		WScript.StdOut.WriteLine "		   --------------------------------------"
		WScript.StdOut.Write	 "		          Your Selection (A-Q):" 
		
		input = Left(UCase(Trim(WScript.StdIn.ReadLine)),1)
		
		formatScreen(2)
		
		Select Case input
			Case "A"
				Call DesktopManagement
			Case "B"
				Call SystemManagement
			Case "C"
				Call UserManagement
			Case "D"
				Call DiskManagement
			Case "E"
				Call Utilities
			Case "Q"
				Call Quit
			Case Else
				WScript.StdOut.WriteLine "ERROR: Enter one of the letters in the menu above."
				WScript.StdOut.Write "Press ENTER to continue..."
				input = WScript.StdIn.ReadLine
		End Select
	Loop
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub DesktopManagement
	Dim input
	Do
		formatScreen(16)
		
		WScript.StdOut.WriteLine "		     	   	Desktop Management"
		WScript.StdOut.WriteLine "		   --------------------------------------"
		WScript.StdOut.WriteLine "		         1-Desktop Settings"
		WScript.StdOut.WriteLine "		         2-Shortcuts"
		WScript.StdOut.WriteLine "		         3-Events"
		WScript.StdOut.WriteLine "		         4-Scheduled Tasks"
		WScript.StdOut.WriteLine "		         5-Return to Main Menu"
		WScript.StdOut.WriteLine "		   --------------------------------------"
		WScript.StdOut.Write	 "		          Your Selection (1-5):" 
		
		input = Left(Trim(WScript.StdIn.ReadLine),1)
		
		formatScreen(2)
		
		Select Case input
			Case "1"
				DesktopSettings 
			Case "2"
				shortCuts
			Case "3"
				
			Case "4"
				
			Case "5"
				Exit Do
			Case Else
				WScript.StdOut.WriteLine "ERROR: Enter one of the numbers in the menu above."
				WScript.StdOut.Write "Press ENTER to continue..."
				input = WScript.StdIn.ReadLine
		End Select
	Loop
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub SystemManagement
	Dim input
	Do
		formatScreen(16)
		
		WScript.StdOut.WriteLine "		     	   	System Management"
		WScript.StdOut.WriteLine "		   --------------------------------------"
		WScript.StdOut.WriteLine "		         1-Restart Computer"
		WScript.StdOut.WriteLine "		         2-Login Script"
		WScript.StdOut.WriteLine "		         3-Logout Script"
		WScript.StdOut.WriteLine "		         4-Shutdown Computer"
		WScript.StdOut.WriteLine "		         5-Return to Main Menu"
		WScript.StdOut.WriteLine "		   --------------------------------------"
		WScript.StdOut.Write	 "		          Your Selection (1-5):" 
		
		input = Left(Trim(WScript.StdIn.ReadLine),1)
		
		formatScreen(2)
		
		Select Case input
			Case "1"
				
			Case "2"
				
			Case "3"
				
			Case "4"
				
			Case "5"
				Exit Do
			Case Else
				WScript.StdOut.WriteLine "ERROR: Enter one of the numbers in the menu above."
				WScript.StdOut.Write "Press ENTER to continue..."
				input = WScript.StdIn.ReadLine
		End Select
	Loop
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub UserManagement
	Dim input
	Do
		formatScreen(15)
		
		WScript.StdOut.WriteLine "		     	   	User Management"
		WScript.StdOut.WriteLine "		   --------------------------------------"
		WScript.StdOut.WriteLine "		         1-Create A Group"
		WScript.StdOut.WriteLine "		         2-Create Users"
		WScript.StdOut.WriteLine "		         3-List All Users"
		WScript.StdOut.WriteLine "		         4-List All Groups"
		WScript.StdOut.WriteLine "		         5-Delete A User"
		WScript.StdOut.WriteLine "		         6-Return to Main Menu"
		WScript.StdOut.WriteLine "		   --------------------------------------"
		WScript.StdOut.Write	 "		          Your Selection (1-6):" 
		
		input = Left(Trim(WScript.StdIn.ReadLine),1)
		
		formatScreen(2)
		
		Select Case input
			Case "1"
				
			Case "2"
				
			Case "3"
				
			Case "4"
				
			Case "5"
			
			Case "6"
				Exit Do
			Case Else
				WScript.StdOut.WriteLine "ERROR: Enter one of the numbers in the menu above."
				WScript.StdOut.Write "Press ENTER to continue..."
				input = WScript.StdIn.ReadLine
		End Select
	Loop
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub DiskManagement
	Dim input
	Do
		formatScreen(18)
		
		WScript.StdOut.WriteLine "		     	   File System Management"
		WScript.StdOut.WriteLine "		   --------------------------------------"
		WScript.StdOut.WriteLine "		         1-Drives"
		WScript.StdOut.WriteLine "		         2-Folders"
		WScript.StdOut.WriteLine "		         3-Return to Main Menu"
		WScript.StdOut.WriteLine "		   --------------------------------------"
		WScript.StdOut.Write	 "		          Your Selection (1-3):" 
		
		input = Left(Trim(WScript.StdIn.ReadLine),1)
		
		formatScreen(2)
		
		Select Case input
			Case "1"
				
			Case "2"
				
			Case "3"
				Exit Do
			Case Else
				WScript.StdOut.WriteLine "ERROR: Enter one of the numbers in the menu above."
				WScript.StdOut.Write "Press ENTER to continue..."
				input = WScript.StdIn.ReadLine
		End Select
	Loop
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub Utilities
	Dim input
	Do
		formatScreen(19)
		
		WScript.StdOut.WriteLine "		     	   Miscellaneous Tasks"
		WScript.StdOut.WriteLine "		   --------------------------------------"
		WScript.StdOut.WriteLine "		         1-Software"
		WScript.StdOut.WriteLine "		         2-Return to Main Menu"
		WScript.StdOut.WriteLine "		   --------------------------------------"
		WScript.StdOut.Write	 "		          Your Selection (1-2):" 
		
		input = Left(Trim(WScript.StdIn.ReadLine),1)
		
		formatScreen(2)
		
		Select Case input
			Case "1"
				
			Case "2"
				Exit Do
			Case Else
				WScript.StdOut.WriteLine "ERROR: Enter one of the numbers in the menu above."
				WScript.StdOut.Write "Press ENTER to continue..."
				input = WScript.StdIn.ReadLine
		End Select
	Loop
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub Quit
	Dim logoutMessages
	logoutMessages = Array("Goodbye!", "Adios!", "We hope you enjoyed the ride, please come again.", "Be seein' you.", "Smell ya later.", "Thank you come again.", "Need more messages here.", "I'll be back!")
	WScript.StdOut.Write logoutMessages(randomNumber(UBound(logoutMessages)))
	WScript.Sleep(2000)
	WScript.Quit
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub DesktopSettings

	dim objshell
	Dim userName
	Dim ObjFileSys
	Dim windowsDir
	Dim wallpaper
	Dim RegKeyMain
	Dim RegScreenSaver 

	'change wallpaper	
	Set objShell = WScript.CreateObject("WScript.Shell")
	userName = objShell.ExpandEnvironmentStrings("%USERNAME%")
	Set ObjFileSys = CreateObject("Scripting.FileSystemObject")
	
	windowsDir = ObjFileSys.GetSpecialFolder(0)
	wallpaper = "C:\Windows\Web\Wallpaper\Windows\img0.jpg"
	
	objShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaper", wallpaper
	objShell.Run "c:\windows\System32\RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters", 1, True

	'change screensaver settings
	
	RegKeyMain = "HKEY_CURRENT_USER\Control Panel\Desktop\"
	RegScreenSaver = "C:\WINDOWS\system32\logon.scr"
	
  	objshell.RegWrite RegKeyMain & "ScreenSaveActive", 1, "REG_SZ"
    objshell.RegWrite RegKeyMain & "ScreenSaverIsSecure", 1, "REG_SZ"
    objshell.RegWrite RegKeyMain & "ScreenSaveTimeOut", 300, "REG_SZ"
    objShell.RegWrite RegKeyMain & "SCRNSAVE.EXE", RegScreenSaver, "REG_SZ"

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub shortCuts

	Dim objShell
	Dim strDesktop
	Dim objShortcut
	Dim objUrlLink
	Dim userURL
	Dim shortcutName
	Dim objFileSystem
	Dim confirm

	
	Set objShell = CreateObject("Wscript.shell")
	Set objFileSystem = CreateObject("Scripting.FileSystemObject")
	strDesktop = objShell.SpecialFolders("Desktop")
	Set objShortcut = objShell.CreateShortcut(strDesktop + "\Task List.lnk")
	objShortcut.IconLocation = "C:\myicon.ico"
	objShortcut.TargetPath = "%windir%\system32\cmd.exe"
	objShortcut.Arguments = "/k tasklist"
	objShortCut.Hotkey = "ALT+CTRL+T"
	objShortCut.Save	
	

	Do
	
		WScript.StdOut.Write "Please Enter the website URL for your shortcut: "
		userURL = WScript.StdIn.ReadLine 
		
		If validwebpage(userURL) Then
			Exit Do
		Else
			WScript.StdOut.WriteLine "Error: webpage must be in http://(sitename).(domain)"
		End if
	Loop
	
	Set objUrlLink = objShell.CreateShortcut(strDesktop+"\Link to user webpage.URL")
	objUrlLink.TargetPath = userURL
	objUrlLink.Save
	
	WScript.StdOut.Write "Please Enter the name of the shortcut you would like to delete: "
	shortcutName = strDesktop & "\" & WScript.StdIn.ReadLine & ".lnk"
		Do
	
		WScript.StdOut.Write "Are you sure you would like to remove the shortcut? (y/n):"
		confirm = UCase(WScript.StdIn.ReadLine)
		

		
		If confirm = "Y" Then
			WScript.StdOut.WriteLine "Deleting shortcut..."
			'deletes shortcut
			objFileSystem.DeleteFile(shortcutName)
			Exit do
		ElseIf confirm = "N" Then
			WScript.StdOut.WriteLine "shortcut will not be deleted"
			Exit Do
		End if
		
		
	Loop
	
	


End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'====================================================================================================================================================================================================
'====================================================================================================================================================================================================

Function isCScript()
	If LCase( Right( WScript.FullName, 12 ) ) = "\cscript.exe" Then
		isCScript = True
	Else
		isCScript = False
	End If
End Function

'25 lines on standard cmd
Function formatScreen(lines)
	Dim counter
	For counter = 1 To lines
		WScript.StdOut.WriteLine
	Next
End Function

Function randomNumber(max)
	Randomize
	randomNumber = Int((max + 1) * Rnd)
End Function

Function validwebpage(input)

	Dim regex

	Set regex = New RegExp
	
	With regex
		.Pattern = "^http\://[a-zA-Z0-9\-\.]+\.[a-zA-Z]{2,3}(/\S*)?$"
		.IgnoreCase = False
		.Global = False
	End With
	
	If regex.Test(input) Then
		validwebpage = True
	Else
		validwebpage = False
	End if

End Function