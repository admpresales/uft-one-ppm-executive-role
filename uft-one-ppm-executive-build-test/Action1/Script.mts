'===========================================================================================
'20200929 - DJ: Updated the step to click the Done button when looking into changing from the calculated risk to the 
'			override value
'20200929 - DJ: Updated improper syntax on the loop exit
'20200929 - DJ: Added .sync statements after .click statements and additional tuning
'20200929 - DJ: Added sync loop for clicking down arrow for risk override
'20201001 - DJ: Added ClickLoop function to leverage it, removed duplicative code
'			Added traditional OR click to force autoscroll if the resolution is too low on the UFT machine to have
'			the 2nd verified dashboard to be displayed, plus changed the selection of the hamburger_menu to be VRI off
'			of the dashboard title.
'20201001 - DJ: Updated manual reporter event error handling
'20201001 - DJ: Updated manual reporter event error handling, shortened the text find for efficiency
'20201008 - DJ: Updated to use traditional OR for Budget by Business Objective Hamburger Menu as on some lower resolutions, the
'				hamburger_menu object isn't being seen by the AI engine.  Additionally, you can have the hamburger menu item displayed
'				but when clicked, the Maximize could be off of the screen.  The Maximize will be able to be changed back to AI when
'				autoscroll option is added in a future version of UFT One.  Lastly, the down triangle at the end, depending on the 
'				resolution might, or might not be the lowest visible down triangle, replaced with traditional OR.
'20201020 - DJ: Updated to handle changes coming in UFT One 15.0.2
'				Commented out the msgbox, which can cause UFT One to be in a locked state when executed from Jenkins
'20201022 - DJ: Updated ClickLoop to gracefully abort if failure number reached
'				Updated failure abort to be 3 instead of 90
'20210208 - DJ: You can use the public PPM demo http://ppmdemo.mfadvantageinc.com/menu.html and this will work, edit data table if you want to run against nimbusserver
'				Added logic to enumerate handled browsers, script as is will fail on 2nd iteration on purpose
'				Updated to exclusively use AI-based object recognition, script will no longer function in 15.0.1
'				Working with R&D on autoscroll issue, uncomment the commented code if it fails.
'===========================================================================================

Public Function Logout
	'===========================================================================================
	'BP:  Logout
	'===========================================================================================
	AIUtil.RunSettings.AutoScroll.Enable "up", 10
	If AIUtil("profile").Exist(0) Then
		AIUtil("profile").Click
	Else
		AIUtil("profile", micAnyText, micFromTop, 1).Click
	End If
	AIUtil.RunSettings.AutoScroll.Enable "down", 2
	AppContext.Sync																				'Wait for the browser to stop spinning
	If AIUtil.FindText("Sign Out").Exist(0) Then
		AIUtil.FindText("Sign Out").Click
	Else
		AIUtil.FindText("Sign Out", micFromTop, 1).Click
	End If
	AppContext.Sync																				'Wait for the browser to stop spinning
	
	AppContext.Close																			'Close the application at the end of your script

End Function

Function ClickLoop (AppContext, ClickStatement, SuccessStatement)
	
	Dim Counter
	
	Counter = 0
	Do
		ClickStatement.Click
		AppContext.Sync																				'Wait for the browser to stop spinning
		Counter = Counter + 1
		wait(1)
		If Counter >=3 Then
			Reporter.ReportEvent micFail, "Click Statement", "The Success Statement '" & SuccessStatement & "' didn't display within " & Counter & " attempts.  Aborting action"
			Logout
			ExitAction
		End If
	Loop Until SuccessStatement.Exist(10)
	AppContext.Sync																				'Wait for the browser to stop spinning

End Function


Dim BrowserExecutable, Counter, rc

While Browser("CreationTime:=0").Exist(0)   												'Loop to close all open browsers
	Browser("CreationTime:=0").Close 
Wend

Select Case DataTable.Value("BrowserName")
	Case "IEXPLORE"
		Reporter.ReportEvent micPass, "Browser Support", "The browser '" & DataTable.Value("BrowserName") & "' is supported, proceeding."
	Case "CHROME"
		Reporter.ReportEvent micPass, "Browser Support", "The browser '" & DataTable.Value("BrowserName") & "' is supported, proceeding."
	Case "FIREFOX"
		Reporter.ReportEvent micPass, "Browser Support", "The browser '" & DataTable.Value("BrowserName") & "' is supported, proceeding."
	Case Else
		Reporter.ReportEvent micFail, "Browser Support", "The browser '" & DataTable.Value("BrowserName") & "' is not supported, aborting action."
		ExitAction
End Select

BrowserExecutable = DataTable.Value("BrowserName") & ".exe"
SystemUtil.Run BrowserExecutable,"","","",3													'launch the browser specified in the data table
Set AppContext=Browser("CreationTime:=0")													'Set the variable for what application (in this case the browser) we are acting upon

'===========================================================================================
'BP:  Navigate to the PPM Launch Pages
'===========================================================================================

AppContext.ClearCache																		'Clear the browser cache to ensure you're getting the latest forms from the application
AppContext.Navigate DataTable.Value("URL")													'Navigate to the application URL
AppContext.Maximize																			'Maximize the application to give the best chance that the fields will be visible on the screen
AppContext.Sync																				'Wait for the browser to stop spinning
AIUtil.SetContext AppContext																'Tell the AI engine to point at the application

'===========================================================================================
'BP:  Click the Executive Overview link
'===========================================================================================
AIUtil.FindText("Executive Overview").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Ron Steel (CIO) link to launch PPM as Ron Steel
'===========================================================================================
AIUtil.FindTextBlock("Ron Steel").Click
AppContext.Sync	
																			'Wait for the browser to stop spinning
If AIUtil.FindText("bubble").Exist(20) Then
	Reporter.ReportEvent micPass, "Find the bubble text", "The text did display within the default .Exist timeout, this means that the Porfolio Dashboard portlet didn't load"
Else 
	Reporter.ReportEvent micFail, "Find the bubble text", "The text didn't display within the default .Exist timeout, this means that the Porfolio Dashboard portlet didn't load"
End If

'===========================================================================================
'BP:  Hover over each Business Objective category to capture the changes in the Porfolio Scorecard
'===========================================================================================
AIUtil.FindTextBlock("Regulatory Compliance").Hover
AIUtil.FindTextBlock("9 Month Release Cycle").Hover
AIUtil.FindTextBlock("Reduce Customer Churn").Hover
AIUtil.FindTextBlock("10% Increase in Revenue").Hover
AIUtil.FindTextBlock("15% Growth in Partner Channels").Hover
AIUtil.FindTextBlock("Cost Containment").Hover

'===========================================================================================
'BP:  Verify that the Budget by Business Objective dashboard element is displayed
'		Added a traditional OR click on the dashboard name to force scroll if the 
'		resolution of the machine is too small to have the hamburger menu be displayed
'===========================================================================================
'Counter = 0
'While AIUtil("hamburger_menu", micAnyText, micWithAnchorOnLeft, AIUtil.FindTextBlock("Budget by Business Objective (This Year)")).Exist(0) = FALSE
'	Counter = Counter + 1
'	wait(1)
'	If Counter >=3 Then
'		Reporter.ReportEvent micFail, "Click Budget by Business Objective (This Year) Button", "The button didn't display within " & Counter & " attempts.  Aborting run."
'		AIUtil.SetContext AppContext																'Tell the AI engine to point at the application
'		ExitIteration
'	End  If
'Wend
AIUtil("hamburger_menu", micAnyText, micWithAnchorOnLeft, AIUtil.FindTextBlock("Budget by Business Objective (This Year)")).Click
AppContext.Sync																				'Wait for the browser to stop spinning
'Counter = 0
'While AIUtil.FindTextBlock("Maximize").Exist(0) = FALSE
'	Counter = Counter + 1
'	wait(1)
'	If Counter >=3 Then
'		Reporter.ReportEvent micFail, "Click Maximize Button", "The Maximize text didn't display within " & Counter & " attempts.  Aborting run."
'		AIUtil.SetContext AppContext																'Tell the AI engine to point at the application
'		ExitIteration
'	End  If
'Wend
AIUtil.FindTextBlock("Maximize").Click	
AppContext.Sync

'===========================================================================================
'BP:  Hover over each Business Objective category to capture the changes in the Porfolio Scorecard
'===========================================================================================
AIUtil.FindTextBlock("Regulatory Compliance").Hover
AIUtil.FindTextBlock("Reduce Customer Chum").Hover
AIUtil.FindTextBlock("Cost Containment").Hover
AIUtil.FindTextBlock("9 Month Release Cycle").Hover
AIUtil.FindTextBlock("15% Growth in Partner Channels").Hover
AIUtil.FindTextBlock("10% Increase in Revenue").Hover

'===========================================================================================
'BP:  Search for portfolio (itfm)
'===========================================================================================
AIUtil("search").Search "portfolio (itfm)"
AIUtil.FindTextBlock("Portfolio (ITFM) (DASHBOARD)").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Trial Portfolio to exercise drill down
'===========================================================================================
AIUtil.FindText("Trial Portfolio").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Marketing WebPortaI V2 to exercise drill down to the project dashboard
'===========================================================================================
AIUtil.FindTextBlock("Marketing WebPortaI V2").Click
AppContext.Sync																				'Wait for the browser to stop spinning
rc = AIUtil.FindText("Requirements Analysis").Exist

'===========================================================================================
'BP:  Click the down triangle to show you could override the calculated health
'===========================================================================================
AIUtil("down_triangle", micAnyText, micWithAnchorOnLeft, AIUtil.FindTextBlock(micAnyText, micWithAnchorAbove, AIUtil.FindTextBlock("Calculated health"))).Click

'===========================================================================================
'BP:  Click Done button
'===========================================================================================
Set ClickStatement = AIUtil("button", "Done")
Set SuccessStatement = AIUtil.FindText("Requirements Analysis")
ClickLoop AppContext, ClickStatement, SuccessStatement

'===========================================================================================
'BP:  Logout
'===========================================================================================
AIUtil.RunSettings.AutoScroll.Enable "up", 10
If AIUtil("profile").Exist(0) Then
	AIUtil("profile").Click
Else
	AIUtil("profile", micAnyText, micFromTop, 1).Click
End If
AIUtil.RunSettings.AutoScroll.Enable "down", 2
AppContext.Sync																				'Wait for the browser to stop spinning
If AIUtil.FindText("Sign Out").Exist(0) Then
	AIUtil.FindText("Sign Out").Click
Else
	AIUtil.FindText("Sign Out", micFromTop, 1).Click
End If
AppContext.Sync																				'Wait for the browser to stop spinning

AppContext.Close																			'Close the application at the end of your script

