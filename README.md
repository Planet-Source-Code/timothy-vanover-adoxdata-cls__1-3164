<div align="center">

## AdoxData\.cls


</div>

### Description

This demonstrates how to create a database and components at runtime from a public sub called from the AdoxData class with ADOX 2.1 objects
 
### More Info
 
Call the sub and send it the a string for the database name, and a string for the key table name and one for the detail table name. This will create two tables, with various data types, with a One to many relationship, which will enforce referential integrety.

Make sure to set a project reference to "Ext.2.1 for DDL and Security". Updates can be obtained from Microsoft through "Mdac_typ".


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Timothy Vanover](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/timothy-vanover.md)
**Level**          |Unknown
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/timothy-vanover-adoxdata-cls__1-3164/archive/master.zip)





### Source Code

```
Option Explicit
'* This uses ADOX components to create a database and database
'* objects at runtime. This can be used also to create databases
'* for applications instead of an the actual Microsoft Access
'* application. Set a reference to "Ext.2.1 for DDL and Security"
'* in the project references. Add this class to a project and call
'* CreateAdox passing the Database Name, Table Name, Table Name
'* Submitted by Timothy A. Vanover
'* hdhunter@home.com
Private tbl As ADOX.Table
Private cat As ADOX.Catalog 'the actual database
Private idx As ADOX.Index
Private Pkey As ADOX.Key
Public Sub CreateAdox(strCatalogName As String, _
  strTableNameOne As String, _
  strTableNameTwo As String)
 Set cat = New ADOX.Catalog
 On Error GoTo MyError
'* This creates the actual database.
 cat.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
 App.Path & "\" & strCatalogName & ".mdb"
 Set tbl = New ADOX.Table
 With tbl
 .Name = strTableNameOne
 Set .ParentCatalog = cat
 .Columns.Append "MyPrimaryKey", adInteger 'long data type
 .Columns("MyPrimaryKey").Properties("AutoIncrement") = True 'auto number
 .Columns.Append "MyIntegerData", adSmallInt 'Integer data type
 .Columns.Append "MyStringData", adVarWChar, 25 'string size of 25
 End With
 cat.Tables.Append tbl 'add the table to the database
 Set Pkey = New ADOX.Key 'create new key object
 With Pkey
 .Name = "MyPrimaryKey"
 .Type = adKeyPrimary
 .Columns.Append "MyPrimaryKey"
 End With
 tbl.Keys.Append Pkey
 Set Pkey = Nothing
 Set idx = New ADOX.Index
 With idx
 .Unique = False 'duplicates allowed
 .Name = "MyIntegerData"
 .Columns.Append "MyIntegerData"
 End With
 tbl.Indexes.Append idx
 Set idx = Nothing
 Set idx = New ADOX.Index
 With idx
 .Unique = True 'NO duplicates allowed
 .Name = "MyStringData"
 .Columns.Append "MyStringData"
 End With
 tbl.Indexes.Append idx
 Set idx = Nothing
 Set tbl = Nothing
'* Create a detail Table with a memo Field, and foreign key
 Set tbl = New ADOX.Table
 With tbl
 .Name = strTableNameTwo
 Set .ParentCatalog = cat
 .Columns.Append "MyPrimaryKey", adInteger 'Long data type
 .Columns.Append "MyMemoData", adLongVarWChar 'Memo data type
 End With
 cat.Tables.Append tbl
 Set Pkey = New ADOX.Key
 With Pkey 'set relationship
 .Name = "MyPrimaryKey"
 .Type = adKeyForeign
 .RelatedTable = strTableNameOne
 .Columns.Append "MyPrimaryKey"
 .Columns("MyPrimaryKey").RelatedColumn = "MyPrimaryKey"
 .UpdateRule = adRICascade 'Enforce Referential Integrity
 End With
 tbl.Keys.Append Pkey
 Set tbl = Nothing
 Set Pkey = Nothing
 Set cat = Nothing
 Exit Sub
MyError:
 Debug.Print Err.Number & Space$(1) & Err.Description
End Sub
```

