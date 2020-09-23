<div align="center">

## Retrieve Values from Web Page \(WebBrowser Control\)


</div>

### Description

Allows you to retrieve values from fields (i.e. hidden fields) in a web page loaded using the WebBrowser control. Very simple function, but I've found it extremely useful. Please don't forget to leave a comment to let me know what you think!

To call this function try this example:

MyValue = GetWebValue(MyWebBrowser,"fieldname")

Will return the value or nothing if not found.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Stephen B\. Gauntt](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/stephen-b-gauntt.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/stephen-b-gauntt-retrieve-values-from-web-page-webbrowser-control__1-32629/archive/master.zip)





### Source Code

```
Function GetWebValue(WebBrowser As WebBrowser, value As String)
 'Checking if a Frame Page is being displayed
 On Error GoTo EndStuff:
 Set LkpWeb = WebBrowser.Document
 If LkpWeb.Frames.Length > 0 Then
 'Cycle through the frames
 For i = 0 To LkpWeb.Frames.Length - 1
 Set Lkp = LkpWeb.Frames(i).Document.All
 On Error Resume Next
 GetWebValue = Lkp.Item(CStr(value)).value
 If Err.Number = 0 Then
 Exit For
 End If
 DoEvents
 Next
 Else
 Set LkpWeb = WebBrowser.Document.All
 On Error Resume Next
 GetWebValue = LkpWeb.Item(CStr(value)).value
 End If
EndStuff:
End Function
```

