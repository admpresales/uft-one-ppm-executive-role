'===========================================================================================
'20200929 - DJ: Updated the step to click the Done button when looking into changing from the calculated risk to the 
'			override value
'20200929 - DJ: Updated improper syntax on the loop exit
'20200929 - DJ: Added .sync statements after .click statements and additional tuning
'===========================================================================================

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
AIUtil.FindText("Size of bubble indicates").Exist

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
'===========================================================================================
AIUtil("hamburger_menu", micNoText, micFromBottom, 1).Click
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
AIUtil("down_triangle", micNoText, micFromBottom, 1).Click
AppContext.Sync																				'Wait for the browser to stop spinning
AIUtil.FindTextBlock("Override health").Exist

'===========================================================================================
'BP:  Click Done button
'===========================================================================================
Counter = 0
Do
	AIUtil("button", "Done").Click
	Counter = Counter + 1
	wait(1)
	If Counter >=90 Then
		msgbox("Something is broken, status of the request hasn't shown up to be approved.")
		Reporter.ReportEvent micFail, "Click the Done button", "The Done button click wasn't accepted within " & Counter & " seconds."
		Exit Do
	End If
Loop While AIUtil("button", "Done").Exist(1)
AppContext.Sync																				'Wait for the browser to stop spinning
AIUtil.FindText("Requirements Analysis").Exist

'===========================================================================================
'BP:  Logout.  Use traditional OR
'===========================================================================================
Browser("Project Overview").Page("Project Overview").WebElement("menuUserIcon").Click
AppContext.Sync																				'Wait for the browser to stop spinning
AIUtil.FindTextBlock("Sign Out (Ronald Steel)").Click
AppContext.Sync																				'Wait for the browser to stop spinning

AppContext.Close																			'Close the application at the end of your script

