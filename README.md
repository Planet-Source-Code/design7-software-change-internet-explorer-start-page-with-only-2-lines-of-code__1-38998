<div align="center">

## Change Internet Explorer Start page with only 2 lines of code\!\!\!


</div>

### Description

Change Internet Explorer Start page with only 2 lines of code!!!

Please vote for my code if you like it
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Design7 Software](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/design7-software.md)
**Level**          |Beginner
**User Rating**    |4.1 (29 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VB Script
**Category**       |[Registry](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/registry__1-36.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/design7-software-change-internet-explorer-start-page-with-only-2-lines-of-code__1-38998/archive/master.zip)





### Source Code

```
Set wshshell = CreateObject("WScript.Shell")
wshshell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Start Page", "http://www.newstartpage.com"
```

