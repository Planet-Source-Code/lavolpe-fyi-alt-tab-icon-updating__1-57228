<div align="center">

## FYI: Alt\+Tab Icon Updating


</div>

### Description

How do you change the icon that appears in the Alt+Tab window? This is nothing new can be found on MSDN. The notes I added here are of interest. More info for skinning forms. Changing a form's icon doesn't affect the Alt+Tab window; you need to change the icon in an upper level, hidden window. Code below shows how. Notes to keep in mind.... The icon does NOT have to be one assigned to a form! If not, the only thing to remember is to cache that icon handle and don't destroy it until your application closes or you reassign using the same code below. Should you not cache the icon, and it is destroyed, the Alt+Tab will show a "blank/invisible" icon instead.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[LaVolpe](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lavolpe.md)
**Level**          |Intermediate
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/lavolpe-fyi-alt-tab-icon-updating__1-57228/archive/master.zip)





### Source Code

```
Public Sub SetApplicationIcon _
  (hIcon As Long, mainHwnd As Long)
' hIcon must be a 32x32 pixel icon
' mainHwnd can be any open form
' This will not work in IDE but
' works when compiled
Dim tHwnd As Long, cParent As Long
Const ICON_BIG As Long = 1
Const GWL_HWNDPARENT = (-8)
' Get starting point
tHwnd = GetWindowLong(mainHwnd, GWL_HWNDPARENT)
' Get EXE's wrapper class
' (all VB compiled apps have them)
Do While tHwnd
  cParent = tHwnd
  tHwnd = GetWindowLong(cParent, _
            GWL_HWNDPARENT)
Loop
On Error Resume Next
' tell the wrapper class what icon to use
PostMessage cParent, WM_SETICON, _
   ICON_BIG, ByVal hIcon
End Sub
```

