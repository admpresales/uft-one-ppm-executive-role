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
'===========================================================================================

Function ClickLoop (AppContext, ClickStatement, SuccessStatement)
	
	Dim Counter
	
	Counter = 0
	Do
		ClickStatement.Click
		AppContext.Sync																				'Wait for the browser to stop spinning
		Counter = Counter + 1
		wait(1)
		If Counter >=90 Then
			msgbox("Something is broken, the Requests hasn't shown up")
			Reporter.ReportEvent micFail, "Click the Search text", "The Requests text didn't display within " & Counter & " attempts."
			Exit Do
		End If
	Loop Until SuccessStatement.Exist(1)
	AppContext.Sync																				'Wait for the browser to stop spinning

End Function

Dim BrowserExecutable, Counter

While Browser("CreationTime:=0").Exist(0)   												'Loop to close all open browsers
	Browser("CreationTime:=0").Close 
Wend
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
AppContext.Sync																				'Wait for the browser to stop spinning
If AIUtil.FindText("Size of bubble indicates").Exist Then
	Reporter.ReportEvent micPass, "Find the Size of bubble indicates text", "The text did display within the default .Exist timeout, this means that the Porfolio Dashboard portlet didn't load"
Else 
	Reporter.ReportEvent micFail, "Find the Size of bubble indicates text", "The text didn't display within the default .Exist timeout, this means that the Porfolio Dashboard portlet didn't load"
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
Browser("Project Overview").Page("Dashboard - Overview Dashboard").WebElement("Budget by Business Objective (This Year)").Click
Set TextAnchor = AIUtil.FindText("Budget by Business Objective (This Year)")				'Set the IconAnchor to be the profile icon
Set ValueAnchor = AIUtil("hamburger_menu", micNoText, micWithAnchorOnLeft, TextAnchor)				'Set the Value field to be an "input" field, with any text, with the IconAnchor to its left
ValueAnchor.Click
AppContext.Sync																				'Wait for the browser to stop spinning
AIUtil.FindTextBlock("Maximize").Click
AppContext.Sync																				'Wait for the browser to stop spinning

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
AIUtil.FindText("Requirements Analysis").Exist

'===========================================================================================
'BP:  Click the down triangle to show you could override the calculated health
'===========================================================================================
Set ClickStatement = AIUtil("down_triangle", micNoText, micFromBottom, 1)
Set SuccessStatement = AIUtil.FindTextBlock("Override health")
ClickLoop AppContext, ClickStatement, SuccessStatement

'===========================================================================================
'BP:  Click Done button
'===========================================================================================
Set ClickStatement = AIUtil("button", "Done")
Set SuccessStatement = AIUtil.FindText("Requirements Analysis")
ClickLoop AppContext, ClickStatement, SuccessStatement

'===========================================================================================
'BP:  Logout.  Use traditional OR
'===========================================================================================
Browser("Project Overview").Page("Project Overview").WebElement("menuUserIcon").Click
AppContext.Sync																				'Wait for the browser to stop spinning
AIUtil.FindTextBlock("Sign Out (Ronald Steel)").Click
AppContext.Sync																				'Wait for the browser to stop spinning

AppContext.Close																			'Close the application at the end of your script

