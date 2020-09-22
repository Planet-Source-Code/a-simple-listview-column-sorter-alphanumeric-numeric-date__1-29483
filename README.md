<div align="center">

## A simple ListView Column Sorter \(alphanumeric, numeric, date\)


</div>

### Description

A handy module used for sorting a ListView Column (in report view). The column can be sorted alphanumerically, numerically or by date (ascending and descending)
 
### More Info
 
ListViewControl As MSComctlLib.ListView

ColumnIndex as Integer

SortType As Integer

SortOrder As Integer

The listview date is in the general date format

Will allow zero length strings in the sort for all types.

Will not allow non-numeric values for sortNumeric (except zero length strings)

Will not allow anything other than a date or zero length string for sortDate

True if executed without error

False if executed with error


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Beginner
**User Rating**    |4.1 (58 globes from 14 users)
**Compatibility**  |VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/a-simple-listview-column-sorter-alphanumeric-numeric-date__1-29483/archive/master.zip)





### Source Code

```

Public Const sortAlphanumeric = 0
Public Const sortNumeric = 1
Public Const sortDate = 2
Public Const sortAscending = 3
Public Const sortDescending = 4
Function SortColumn(ByVal ListViewControl As MSComctlLib.ListView, ColumnIndex As Integer, SortType As Integer, SortOrder As Integer) As Boolean
 Dim x As Integer, y As Integer
 On Error GoTo ErrHandler
 Select Case SortType
 '*** Alphanumeric sort
 Case sortAlphanumeric
   DoSort ListViewControl, SortOrder, ColumnIndex - 1
 '*** Numeric Sort
 Case sortNumeric
   Dim strMax As String, strNew As String
   'Find the longest (whole) number string length in the column
   If ColumnIndex > 1 Then
    For x = 1 To ListViewControl.ListItems.Count
     If Len(ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1)) <> 0 Then 'ignores 0 length strings
      If Len(CStr(Int(ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1)))) > Len(strMax) Then
        strMax = CStr(Int(ListViewControl.ListItems(x).SubItems(ColumnIndex - 1)))
      End If
     End If
    Next
   Else
    For x = 1 To ListViewControl.ListItems.Count
     If Len(ListViewControl.ListItems(x)) <> 0 Then
      If Len(CStr(Int(ListViewControl.ListItems(x)))) > Len(strMax) Then
        strMax = CStr(Int(ListViewControl.ListItems(x)))
      End If
     End If
    Next
   End If
   'hide the control - speeds up the sort
   ListViewControl.Visible = False
   If ColumnIndex > 1 Then
    For x = 1 To ListViewControl.ListItems.Count
     If Len(ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1)) = 0 Then
      ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1) = "0" 'make 0 length strings = to "0"
     ElseIf Len(CStr(Int(ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1)))) < Len(strMax) Then
       'prefix all numbers with 0's as required
       strNew = ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1)
       For y = 1 To Len(strMax) - Len(CStr(Int(ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1))))
        strNew = "0" & strNew
       Next
       ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1) = strNew
     End If
    Next
   Else
    For x = 1 To ListViewControl.ListItems.Count
     If Len(ListViewControl.ListItems(x).Text) = 0 Then
      ListViewControl.ListItems(x).Text = "0" 'make 0 length strings = to "0"
     ElseIf Len(CStr(Int(ListViewControl.ListItems(x)))) < Len(strMax) Then
       'prefix all numbers with 0's as required
       strNew = ListViewControl.ListItems(x).Text
       For y = 1 To Len(strMax) - Len(CStr(Int(ListViewControl.ListItems(x))))
        strNew = "0" & strNew
       Next
       ListViewControl.ListItems(x).Text = strNew
     End If
    Next
   End If
   DoSort ListViewControl, SortOrder, ColumnIndex - 1
   If ColumnIndex > 1 Then
    'Remove preceding 0's
    For x = 1 To ListViewControl.ListItems.Count
     ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1) = CDbl(ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1))
     If ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1) = 0 Then ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1) = ""
    Next
   Else
    'Remove preceding 0's
    For x = 1 To ListViewControl.ListItems.Count
     ListViewControl.ListItems(x).Text = CDbl(ListViewControl.ListItems(x).Text)
     If ListViewControl.ListItems(x).Text = 0 Then ListViewControl.ListItems(x).Text = ""
    Next
   End If
   ListViewControl.Visible = True
 '*** Date Sort
 Case sortDate
   ListViewControl.Visible = False
   If ColumnIndex > 1 Then
    'Convert dates to format that can be sorted alphanumerically
    For x = 1 To ListViewControl.ListItems.Count
     ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1) = Format(ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1), "YYYY MM DD hh:mm:ss")
    Next
    DoSort ListViewControl, SortOrder, ColumnIndex - 1
    'Convert dates back to General Date format
    For x = 1 To ListViewControl.ListItems.Count
     ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1) = Format(ListViewControl.ListItems(x).ListSubItems(ColumnIndex - 1), "General Date")
    Next
   Else
    'Convert dates to format that can be sorted alphanumerically
    For x = 1 To ListViewControl.ListItems.Count
     ListViewControl.ListItems(x).Text = Format(ListViewControl.ListItems(x).Text, "YYYY MM DD hh:mm:ss")
    Next
    DoSort ListViewControl, SortOrder, ColumnIndex - 1
    'Convert dates back to General Date format
    For x = 1 To ListViewControl.ListItems.Count
     ListViewControl.ListItems(x).Text = Format(ListViewControl.ListItems(x).Text, "General Date")
    Next
   End If
   ListViewControl.Visible = True
 End Select
 SortColumn = True
Exit_Function:
 Exit Function
ErrHandler:
 MsgBox Err.Description & " (" & Err.Number & ")", vbOKOnly + vbCritical, "ListView Sort module Error"
 SortColumn = False
 Resume Exit_Function
End Function
Private Sub DoSort(ByVal ListViewControl As MSComctlLib.ListView, SortOrder As Integer, SortKey As Integer)
  If SortOrder = sortAscending Then
   ListViewControl.SortOrder = lvwAscending
  ElseIf SortOrder = sortDescending Then
   ListViewControl.SortOrder = lvwDescending
  End If
  ListViewControl.SortKey = SortKey
  ListViewControl.Sorted = True
End Sub
'******************************************************************
'************** EXAMPLE CALL FROM FORM - ON LISTVIEW COLUMN CLICK
'******************************************************************
'Private Sub lv_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'
' Select Case ColumnHeader.Index
'  Case 1
'    If lv(Index).ColumnHeaders(ColumnHeader.Index).Icon = "up" Then
'     SortColumn lv(Index), ColumnHeader.Index, sortAlphanumeric, sortDescending
'     lv(Index).ColumnHeaders(ColumnHeader.Index).Icon = "down"
'    Else
'     SortColumn lv(Index), ColumnHeader.Index, sortAlphanumeric, sortAscending
'     lv(Index).ColumnHeaders(ColumnHeader.Index).Icon = "up"
'    End If
'
'  Case 2
'    If lv(Index).ColumnHeaders(ColumnHeader.Index).Icon = "up" Then
'     SortColumn lv(Index), ColumnHeader.Index, sortNumeric, sortDescending
'     lv(Index).ColumnHeaders(ColumnHeader.Index).Icon = "down"
'    Else
'     SortColumn lv(Index), ColumnHeader.Index, sortNumeric, sortAscending
'     lv(Index).ColumnHeaders(ColumnHeader.Index).Icon = "up"
'    End If
'
'  Case 3
'    If lv(Index).ColumnHeaders(ColumnHeader.Index).Icon = "up" Then
'     SortColumn lv(Index), ColumnHeader.Index, sortDate, sortDescending
'     lv(Index).ColumnHeaders(ColumnHeader.Index).Icon = "down"
'    Else
'     SortColumn lv(Index), ColumnHeader.Index, sortDate, sortAscending
'     lv(Index).ColumnHeaders(ColumnHeader.Index).Icon = "up"
'    End If
'
'
' End Select
'
' For x = 1 To lv(Index).ColumnHeaders.Count
'  If x <> ColumnHeader.Index Then
'   lv(Index).ColumnHeaders(x).Icon = "dot"
'  End If
' Next
'End Sub
```

