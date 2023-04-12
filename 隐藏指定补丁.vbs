' =======================================================================================================
' Version: 1.2.20230323_2117
' Why-cn rewrited this for disabling outdated/unnecessary hotfixs for Windows 8.1 final installation ISO.
' Original content: https://superuser.com/questions/722667/how-to-hide-updates-in-windows-updates-without-gui
' Please read the following contents if you want to reuse, or republish to public sites.
' =======================================================================================================
' Original content: http://www.msfn.org/board/topic/163162-hide-bing-desktop-and-other-windows-updates/
' Maxpsoft May 30, 2013, 9:34:15 PM
' 06/18/2013 Add extra for Bing Desktop v1.3
' 06/28/2013 Updated to continue searching as long as it is finding something otherwise Quit
'
' Original Mike.Moore Dec 17, 2012 on answers.microsoft but when ran it Hide everything so no good.
' Link to script: http://www.msfn.org/board/topic/163162-hide-bing-desktop-and-other-windows-updates/
' You may freely use this script as long as you copy it complete and it remains the same except for adjusting hideupdates.
' If I need to change something then let me know so all may benefit.
' =======================================================================================================

Dim WSHShell, StartTime, ElapsedTime, strUpdateName, strAllHidden
Dim Checkagain 'Find more keep going otherwise Quit

Dim hideupdates(39) 'FOR REUSER: change the number inside the brackets to fit the count - 1 of the hotfixs you want to hide
' for example, if you want to hide 40 hotfixs, the number inside the brackets should be 39.
' FOR REUSER: the contents inside the quote below can be keys inside the hotfixs descriptions, like if you do not want some updates for Microsoft Edge, put "Edge" inside.

hideupdates(0) = "KB2931358" 
hideupdates(1) = "KB2931366" 
hideupdates(2) = "KB2937220" 
hideupdates(3) = "KB2938772"
hideupdates(4) = "KB2949621"
hideupdates(5) = "KB2939153"
hideupdates(6) = "KB2958262"
hideupdates(7) = "KB2958263"
hideupdates(8) = "KB2961072"
hideupdates(9) = "KB2962140"
hideupdates(10) = "KB2965142"
hideupdates(11) = "KB2965500"
hideupdates(12) = "KB2966407"
hideupdates(13) = "KB2968599"
hideupdates(14) = "KB2971203"
hideupdates(15) = "KB2973351"
hideupdates(16) = "KB2975061"
hideupdates(17) = "KB2976627"
hideupdates(18) = "KB2977629"
hideupdates(19) = "KB2978002"
hideupdates(20) = "KB2981580"
hideupdates(21) = "KB2987107"
hideupdates(22) = "KB2989647"
hideupdates(23) = "KB2989930"
hideupdates(24) = "KB2998527"
hideupdates(25) = "KB3001237"
hideupdates(26) = "KB3003057"
hideupdates(27) = "KB3004545"
hideupdates(28) = "KB3013531"
hideupdates(29) = "KB3014442"
hideupdates(30) = "KB3024751"
hideupdates(31) = "KB3038002"
hideupdates(32) = "KB3044374"
hideupdates(33) = "KB3045719"
hideupdates(34) = "KB3046480"
hideupdates(35) = "KB3046737"
hideupdates(36) = "KB3062760"
hideupdates(37) = "KB3133690"
hideupdates(38) = "KB3184143"
hideupdates(39) = "KB4486107"


Set WSHShell = CreateObject("WScript.Shell")

StartTime = Timer 'Start the Timer

Set updateSession = CreateObject("Microsoft.Update.Session")
updateSession.ClientApplicationID = "MSDN Sample Script"
Set updateSearcher = updateSession.CreateUpdateSearcher()
Set searchResult = updateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")

Checkagain = "True"

For K = 0 To 10 'Try 10 times for hide all hotfixs containing the "key" above
If Checkagain = "True" Then
Checkagain = "False"
CheckUpdates
ParseUpdates
End if
Next

' Dialog after the operations were done
ElapsedTime = Timer - StartTime
strTitle = "Windows Updates Hidden"
strText = "The operations were finished."
strText = strText & vbCrLf & ""
strText = strText & strAllHidden
strText = strText & vbCrLf & ""
strText = strText & vbCrLf & "Total Time : " & ElapsedTime
strText = strText & vbCrLf & "===================="
strText = strText & vbCrLf & "Original From Maxpsoft, edited by Why-cn."
strText = strText & vbCrLf & "You are free to use & copy, but read the comments in the vbs source code first."
intType = vbOkOnly

' Silent just comment these 2 lines with a ' and it will run and quit
Set objWshShell = WScript.CreateObject("WScript.Shell")
intResult = objWshShell.Popup(strText, ,strTitle, intType)

' Open Windows Update after remove the comment '
'WSHShell.Run "%windir%\system32\control.exe /name Microsoft.WindowsUpdate"

Set objWshShell = Nothing
Set WSHShell = Nothing
WScript.Quit


Function ParseUpdates 'cycle through updates
For I = 0 To searchResult.Updates.Count-1
Set update = searchResult.Updates.Item(I)
strUpdateName = update.Title
'WScript.Echo I + 1 & "> " & update.Title
For j = 0 To UBound(hideupdates)
if instr(1, strUpdateName, hideupdates(j), vbTextCompare) = 0 then
Else
strAllHidden = strAllHidden _
& vbcrlf & update.Title
update.IsHidden = True'
Checkagain = "True"
end if
Next
Next
End Function

Function CheckUpdates 'check for new updates cause there can be more than one hotfix containing a certain "key"
Set updateSession = CreateObject("Microsoft.Update.Session")
updateSession.ClientApplicationID = "MSDN Sample Script"
Set updateSearcher = updateSession.CreateUpdateSearcher()
Set searchResult = _
updateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")
End Function