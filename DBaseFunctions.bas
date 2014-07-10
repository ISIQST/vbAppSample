Attribute VB_Name = "DBaseFunctions"
Option Explicit
Function CreateTable(td As TableDef, dbDest As Database)
   Dim tdnew As TableDef
   Dim fld As Field
   
   Set tdnew = dbDest.CreateTableDef(td.name, td.Attributes, td.SourceTableName, td.Connect)
   For Each fld In td.Fields
      If (fld.Attributes And dbSystemField) <> dbSystemField Then
         tdnew.Fields.Append tdnew.CreateField(fld.name, fld.Type, fld.Size)
      
         tdnew.Fields(fld.name).AllowZeroLength = fld.AllowZeroLength
         tdnew.Fields(fld.name).DefaultValue = fld.DefaultValue
         tdnew.Fields(fld.name).OrdinalPosition = fld.OrdinalPosition
      End If
   Next
   dbDest.TableDefs.Append tdnew
End Function

Public Sub AddNewDefaultRecords(rsDest As Recordset)
   Dim i%, j%
   
   On Error GoTo errAddnewDefaultRecords
   
    If rsDest.RecordCount = 0 Then
       rsDest.AddNew
       rsDest.Update
    End If
   
errAddnewDefaultRecords:
   If Err <> 0 Then
         MsgBox CStr(Err) + ":" + Error, vbOKOnly, App.Title
         Resume Next
   End If
End Sub

Function TableExists(td As TableDefs, name$) As Boolean
   Dim i
   On Error Resume Next
   i = td(name).name
   If Err = 0 Then
      TableExists = True
   Else
      TableExists = False
      Err = 0
   End If
End Function

Function SynchronizeDatabases(Source$, Dest$)
   Dim dbSource As Database
   Dim dbDest As Database
   Dim tdSource As TableDefs
   Dim tdDest As TableDefs
   Dim rsSource As Recordset
   Dim rsDest As Recordset
   Dim td As TableDef
   Dim fld As Field
   Dim tdnew As TableDef
   Dim i%, j%
   
   On Error GoTo errorhandler
   'open source
   Set dbSource = OpenDatabase(Source)
   Set tdSource = dbSource.TableDefs
   'open dest
   Set dbDest = OpenDatabase(Dest)
   Set tdDest = dbDest.TableDefs
   'for each table in source
   For Each td In tdSource
      'if table is not a system or hidden table
      If (td.Attributes And &H80000000) <> &H80000000 And (td.Attributes And &H2) <> &H2 And (td.Attributes And dbHiddenObject) <> dbHiddenObject Then 'and td.Name <> "PreampChips"
         'if table does not exist in dest
         If Not TableExists(tdDest, td.name) Then
            'add table to dest
            'MsgBox "Add " & td.Name
            
            Call CreateTable(td, dbDest)
            
         Else
            Call AddFieldstoTable(td, tdDest(td.name), dbSource, dbDest)
         End If
         'if source has more than one record, overwrite data in dest
         Call AddRecordstoTable(dbSource, dbDest, td.name)
      End If
   Next
   
errorhandler:

End Function

Function AddFieldstoTable(tdSource As TableDef, tdDest As TableDef, dbSource As Database, dbDest As Database)
   Dim rsSource As Recordset, rsDest As Recordset
   Dim fld As Field
   Dim i%
   
   
   'for each field in source tabledef
   For Each fld In tdSource.Fields
      'if field does not exist in dest
      If Not FieldExists(tdDest, fld.name) Then
         If (fld.Attributes And dbSystemField) <> dbSystemField Then
            'add field to dest tabledef
            tdDest.Fields.Append tdDest.CreateField(fld.name, fld.Type, fld.Size)
            tdDest.Fields(fld.name).AllowZeroLength = fld.AllowZeroLength
            tdDest.Fields(fld.name).DefaultValue = fld.DefaultValue
            tdDest.Fields(fld.name).OrdinalPosition = fld.OrdinalPosition
            'Set rsSource = dbSource.OpenRecordset("Select * From " & tdSource.name)
            Set rsDest = dbDest.OpenRecordset("Select * From " & tdSource.name)
            If rsDest.RecordCount > 0 Then
               rsDest.MoveLast
               rsDest.MoveFirst
               For i = 1 To rsDest.RecordCount
                  rsDest.Edit
                  If tdSource.Fields(fld.name).Type <> 1 Then 'type 1 = Yes/No
                     If tdSource.Fields(fld.name).Type = 10 And tdSource.Fields(fld.name).DefaultValue <> "" Then 'type 10 = string
                        rsDest.Fields(fld.name) = RemoveOutsideQuotes(tdSource.Fields(fld.name).DefaultValue) 'set value to default
                     Else
                        rsDest.Fields(fld.name) = tdSource.Fields(fld.name).DefaultValue 'set value to default
                     End If
                  Else
                     If tdSource.Fields(fld.name).DefaultValue = "No" Then
                        rsDest.Fields(fld.name) = False
                     Else
                        rsDest.Fields(fld.name) = True
                     End If
                  End If
                  rsDest.Update
                  rsDest.MoveNext
               Next
            End If
            'rsSource.Close
            rsDest.Close
         End If
      End If
   Next
End Function

Function AddRecordstoTable(dbSource As Database, dbDest As Database, TableName As String)
   Dim i%, j%
   Dim rsSource As Recordset, rsDest As Recordset
   Set rsSource = dbSource.OpenRecordset("Select * from " & TableName)
    If rsSource.RecordCount > 0 Then
       rsSource.MoveLast
       rsSource.MoveFirst
       'if source has more than one record, overwrite data in dest
       If rsSource.RecordCount > 1 Then
          dbDest.Execute "DELETE * FROM " & TableName
          Set rsDest = dbDest.OpenRecordset("Select * From " & TableName)
          For j = 1 To rsSource.RecordCount
             rsDest.AddNew
             For i = 0 To rsDest.Fields.Count - 1
                If (rsDest.Fields(i).Attributes And dbSystemField) <> dbSystemField Then rsDest.Fields(i) = rsSource.Fields(rsDest.Fields(i).name)
             Next
             rsDest.Update
             rsSource.MoveNext
          Next
       End If
    End If
End Function

Function FieldExists(td As TableDef, name$) As Boolean
   Dim i
   On Error Resume Next
   i = td.Fields(name).name
   If Err = 0 Then
      FieldExists = True
   Else
      FieldExists = False
      Err = 0
   End If
End Function

Function RemoveOutsideQuotes(ByVal locstr$) As String
   If Left(locstr, 1) = Chr(34) And Right(locstr, 1) = Chr(34) Then
      RemoveOutsideQuotes = Mid(locstr, 2, Len(locstr) - 2)
   Else
      RemoveOutsideQuotes = locstr
   End If
End Function
