<div align="center">

## Get current Outlook user


</div>

### Description

Retrieve the user currently logged on in the existing Microsoft Outlook session.
 
### More Info
 
Open a new VB project. Add a button to the form. Add a reference to "Microsoft Outlook ...". Paste the code into the form. Run the project.

Returns the Name property of the recipient object.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Klemens Schmid](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/klemens-schmid.md)
**Level**          |Beginner
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/klemens-schmid-get-current-outlook-user__1-6330/archive/master.zip)





### Source Code

```
Private Sub cmdUser_Click()
'get the user from the current Outlook session
Dim ol As Outlook.Application
Dim ns As NameSpace
Dim oRec As Recipient
Set ol = New Outlook.Application
Set ns = ol.GetNamespace("MAPI")
Call ns.Logon(, , , False)
Set oRec = ns.CurrentUser
MsgBox oRec.Name
End Sub
```

