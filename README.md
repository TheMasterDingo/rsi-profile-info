This is a simple script that queries the RSI website and shows formatted information regarding searched player.  
![](https://i.imgur.com/cbQbmcN.png)
**Install**  
-You will need the HTMLAgilityPack to run it.  
-Easiest way is to download it here https://www.nuget.org/packages/HtmlAgilityPack then 'Download package' 
-Once downloaded extract the .nupkg (with 7zip or other) and navigate to \lib\netstandard2.0 and extract HtmlAgilityPack.dll in a folder of your choice.  
-Open the script with Notepad and paste the path of the .dll into the first line $htmlAgilityPackPath = "PATH TO YOUR DLL"  

**Run**  
-You can either type the username right after the script like ./rsi-profile.ps1 UserName or just run the script and it will ask you for a username  
--Bonus: You can type '-b' to print the Bio information also
**PROFIT**  
