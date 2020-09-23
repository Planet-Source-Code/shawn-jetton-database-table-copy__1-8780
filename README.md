<div align="center">

## Database Table Copy


</div>

### Description

This code will copy all fields from all tables in a database and add them to another identical database.

My company has Rep submitting database reports all the time. I need to append the data from the reports database to a live database. At first I was running querys I had made in the database but our file structure changed and that was no long an option. This is about as simple as it gets but someone might find it useful.
 
### More Info
 
Call the function like this.

AddData "C:\My Documents\Test\From.mdb", "C:\My Documents\Test\To.mdb"

None that I know of.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Shawn Jetton](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/shawn-jetton.md)
**Level**          |Beginner
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/shawn-jetton-database-table-copy__1-8780/archive/master.zip)





### Source Code

```
Public Sub AddData(DataFrom As String, DataTo As String)
Dim dbFrom, dbTo As Database
Dim rsFrom, rsTo As Recordset
Set dbFrom = OpenDatabase(DataFrom)
Set dbTo = OpenDatabase(DataTo)
For n = 0 To dbFrom.TableDefs.Count - 1
    'This search out on table in your database
    If dbFrom.TableDefs(n).Attributes = 0 Then
      Set rsFrom = dbFrom.OpenRecordset(dbFrom.TableDefs(n).Name)
      Set rsTo = dbTo.OpenRecordset(dbTo.TableDefs(n).Name)
    End If
    'Loops through all fields in table and copies from dbFrom to dbTo.
    Do Until rsFrom.EOF
      rsTo.AddNew
      For i = 1 To rsTo.Fields.Count - 1
        If rsFrom.Fields(i) = "" Then GoTo hell
        rsTo.Fields(i) = rsFrom.Fields(i)
hell:
      Next i
      'This updates and moves to the next record in the from database
      rsTo.Update
      rsFrom.MoveNext
    Loop
Next n
dbFrom.Close
dbTo.Close
End Sub
```

