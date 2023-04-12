# Hide-specified-patch-script
## A vbs script designated to hide Windows patches/hotfixes with specified KB no. (Tested on Windows 7~8.1)

Most of its content is from this webpage: [How to hide updates in Windows Updates without GUI - Super User](https://superuser.com/questions/722667/how-to-hide-updates-in-windows-updates-without-gui), and the answer in that question is based on this webpage: [Hide Bing Desktop and other Windows Updates - Unattended Windows 7/Server 2008R2 - MSFN](https://msfn.org/board/topic/163162-hide-bing-desktop-and-other-windows-updates/)

I have edited the script to fit my own purpose, and fixed some bugs inside it, added some prompt and comments. 
Any one can use this script freely, all the efforts belonged to its original creators. Do remember to quote their efforts.

## How to use
First, you need to implement the .reg file to let the Explorer to allow you running the script as administrator. This is important, because it will need to operate the Windows Update, and that needs the administrator permission, and this selection is not shown on default in the Explorer.  
Second, run the vbs script as administrator. You may need to right click the .vbs file and select "Run the script as administrator".  
I suggest you first go to Windows Update, and scan the hotfixes first. When you do the second step, it will not show any window or prompt when its running, so don't be panic or running multiple times. Wait patiently for 2~20 minutes (depends on how many hotfixes you want to hide), and a final prompt will appear. If all successed, running the script second time, it will not show the hotfixes already hidden.

Have fun using the script and if there's any issue, rise an issue.

Why-cn
