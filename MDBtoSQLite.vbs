''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Simple VB MDB to SQLite                                                      '
'  Copyright 2008 Lokkju, <lokkju@lokkju.com>                                       '
'  Original code (c) <rapto@arrakis.es>,<rotoxl@jazzfree.com> unknown license  '
'   found at http://www.sqlite.org/cvstrac/wiki?p=ConverterTools               '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    This program is free software: you can redistribute it and/or modify      '
'    it under the terms of the GNU General Public License as published by      '
'    the Free Software Foundation, either version 3 of the License, or         '
'    (at your option) any later version.                                       '
'                                                                              '
'    This program is distributed in the hope that it will be useful,           '
'    but WITHOUT ANY WARRANTY; without even the implied warranty of            '
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the             '
'    GNU General Public License for more details.                              '
'                                                                              '
'    You should have received a copy of the GNU General Public License         '
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.     '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                                              '
' Usage:                                                                       '
'    cscript.exe MDBtoSQLite.vbs <mdb file> [second mdb file] [...]            '
'    Example: cscript MDBtoSQLite.vbs table.mdb                                '
'    Adding more arguments will process each given file                        '
'                                                                              '
' Description:                                                                 '
'  Takes in an MDB and produces a text file that can be piped into the sqlite  '
'  command line tool to create a database.                                     '
'                                                                              '
' History:                                                                     '
'  0000-00-00 <rapto@arrakis.es>                                               '
'    modified program into a vbs                                               '
'  0000-00-00 <rotoxl@jazzfree.com>                                            '
'    Changed function getSQLiteFieldType according to dao360 docs to handle    '
'     datatypes properly.                                                      '                                 '
'  2008-01-16 <lokkju@lokkju.com>                                              '
'    Fixed INTEGER field types to handle null data, now will insert 0, used to '
'     break it by inserting nothing.  Also modified to accept database path    '
'     from the command line as an argument.                                    '
'  2009-01-21 <lito@list.ru>                                                   '
'    Added function for exportIndexes - create Indexes (except Primary Key)    '
'    Added NOT NULL and DEFAULT values in CREATE TABLE (Now() = CURRENT_TIMESTAMP)          
'    Fixed CREATE TABLE without Primary Key                                    '
'    Added creating TRIGGERs instead of FOREIGN KEY constraits.                '
'    Added Const dbExportData for genereting scripts without INSERT instructions
'    Added type INTEGER PRIMARY KEY for field type COUNTER
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
If LCase(Right(Wscript.FullName, 11)) = "wscript.exe" Then
    Dim strPath
    Dim strCommand
    Dim objShell
    strPath = Wscript.ScriptFullName
    strCommand = "%comspec% /k cscript  " & Chr(34) & strPath & chr(34)
    Set objShell = CreateObject("Wscript.Shell")
    objShell.Run(strCommand)
    Wscript.Quit
End If

Const dbUseJet = 2

'=========================================================
'Script managment constants
'Export INSERT instructions
Const bExportInsert = False
'Create Triggers
Const bExportTrigger = True

'dao datatypes
Const dbBoolean    = 1   
Const dbByte       = 2   
Const dbInteger    = 3   
Const dbLong       = 4   
Const dbCurrency   = 5   
Const dbSingle     = 6   
Const dbDouble     = 7   
Const dbDate       = 8   
Const dbBinary     = 9   
Const dbText       = 10  
Const dbLongBinary = 11   
Const dbMemo       = 12  
Const dbGUID       = 15  
Const dbBigInt     = 16  
Const dbVarBinary  = 17  
Const dbChar       = 18  
Const dbNumeric    = 19  
Const dbDecimal    = 20  
Const dbFloat      = 21  
Const dbTime       = 22  
Const dbTimeStamp  = 23  

'References constants and field attributs
Const dbRelationDeleteCascade = 4096
Const dbRelationUpdateCascade = 256
Const dbAutoIncrField         = 16

Dim sql_keywords1, sql_keywords2

Private Function pad(ByVal valor,ByVal longitud)
    dim ret
    ret="" & valor
    while len(ret)<longitud
        ret="0" & ret
    wend
    pad=ret
End Function

Private Function fechaCadena(ByVal valor)
    dim ret
    'WScript.Echo "fechaCadena",valor
    ret=pad(year(valor),4) & "-" & pad(month(valor),2) & "-" & pad(day(valor),2)
    ret=ret & " " & pad(hour(valor),2) & ":" & pad(minute(valor),2) & ":" & pad(second(valor),2)
    fechaCadena = ret
End Function

Private Function isSQLiteKeyword(ByVal fieldname ) 
    
    Dim ucase_fieldname 
    Dim reservada 

    ucase_fieldname = UCase(fieldname)
    isSQLiteKeyword = False
    For each reservada in sql_keywords1
        If ucase_fieldname = reservada Then
            isSQLiteKeyword = True
            Exit Function
        End If
    Next 
    For each reservada in sql_keywords2
        If ucase_fieldname = reservada Then
            isSQLiteKeyword = True
            Exit Function
        End If
    Next 
End Function

Private Function sql_name(ByVal name )
    If isSQLiteKeyword(name) Or InStr(name, " ") > 0 Then
        sql_name = "[" & name & "]"
    Else
        sql_name = name
    End If
End Function

Function getSQLiteFieldType(db_field , errtype )
    Select Case db_field.Type
        Case dbBoolean 'Yes/No
            getSQLiteFieldType = "BOOLEAN"
        Case dbByte, dbInteger,dbSingle, dbDouble,dbLong
            getSQLiteFieldType = "INTEGER"
        Case dbNumeric, dbBigInt
            getSQLiteFieldType = "NUMBER"
        Case dbDecimal
            getSQLiteFieldType = "NUMBER(" & db_field.Size & ")"
        
        Case dbGUID 
            getSQLiteFieldType = "VARCHAR2"
        
        Case dbFloat
            getSQLiteFieldType = "FLOAT"
        Case dbCurrency
            getSQLiteFieldType = "NUMBER(32,2)"
        Case dbDate, dbTime, dbTimeStamp
            getSQLiteFieldType = "DATE"
        Case dbText
            getSQLiteFieldType = "VARCHAR2(" & db_field.Size & ")"
        Case dbMemo
            getSQLiteFieldType = "TEXT"
        Case dbChar
            getSQLiteFieldType = "CHAR2(" & db_field.Size & ")"
            
        Case dbBinary, dbVarBinary, dbLongBinary
            If errtype Then
                getSQLiteFieldType = "-- error: Field " & db_field.name & " in table " & db_field.SourceTable & " has field type " & db_field.Type _
                & ". Type has been defined as BINARY, and it's data set NULL."
            Else
                getSQLiteFieldType = "BINARY" 'yet unsupported
            End If
        Case Else 'dont know this one
            if errtype Then
                getSQLiteFieldType = "-- error: Field " & db_field.name & " in table " & db_field.SourceTable & " has field type " & db_field.Type _
                & ". Type is UNKNOWN, set to BINARY, and it's data set NULL."
            Else
                getSQLiteFieldType = "BINARY" 'yet unsupported
            End If
    End Select 
End Function

Private Sub exportDatabaseTable(ByRef db , tabla , ts )
    
    Dim rcrdSet 
    Dim print_string
    Dim columna
    Dim field_type 
    dim v, i
    
    Set rcrdSet = db.OpenRecordset("SELECT * FROM " & tabla.name)
    WScript.Echo "   Export Data..."
    i = 1
    While (Not rcrdSet.EOF)
        print_string = "INSERT INTO " & sql_name(tabla.name) & " VALUES ("
        
        For each columna in rcrdSet.Fields
            field_type = getSQLiteFieldType(columna,False)
            If (InStr(1, field_type, "VARCHAR2") <> 0 Or field_type = "TEXT" ) Then
                v=columna.Value
                If Not IsNull(v) Then
                    v=replace(v, "'", "''")
                    v=replace(v, chr(10), "'||chr(10)||'")
                    v=replace(v, chr(11), "'||chr(11)||'")
                    v=replace(v, chr(12), "'||chr(12)||'")
                    v=replace(v, chr(13), "'||chr(13)||'")
                    v=replace(v, chr(9), "'||chr(9)||'")						
                    print_string = print_string & "'" & v & "', "
                Else
                    print_string = print_string & "NULL, "
                End If
            ElseIf field_type = "DATE" Then
                v=columna.Value
                if isnull(v) then
                    print_string = print_string & "NULL, "
                else
                    print_string = print_string & "DATETIME('"& fechaCadena(v) & "', 'YYYY-MM-DD HH24:MI:SS'), "
                End If
            ElseIf (field_type = "BOOLEAN") Then
                If (columna.Value = True) Then
                    print_string = print_string & "1, "
                Else
                    print_string = print_string & "0, "
                End If
            ElseIf (field_type = "INTEGER") Then
                    Dim intval
                    if (isnull(columna.Value)) Then
                    	intval = CStr(0)
                    Else
	                    intval = CStr(columna.Value)
                   	End If
                    If len(intval) = 0 Then intval = "0"
                    If Left(intval, 1) = "." Then intval = "0" & intval
                    print_string = print_string & intval & ", "
            Else
                If (field_type = "BINARY" Or IsNull(columna.Value)) Then
                    'print_string = print_string & "NULL, "
                Else
                    Dim strval 
                    strval = CStr(columna.Value)
                    If Left(strval, 1) = "." Then strval = "0" & strval
                    print_string = print_string & strval & ", "
                End If
            End If
        Next 
        print_string = Mid(print_string, 1, Len(print_string) - 2)
        print_string = print_string & ");"
        on error resume next
        ts.writeline print_string
        if err then
            WScript.Echo print_string 
            WScript.Echo len(print_string)
            err.clear
        end if
        rcrdSet.MoveNext
        if i mod 10000 = 0 then
          WScript.Echo "   Exported " & i & " rows"
        end if
        i = i + 1
    Wend
    
    rcrdSet.Close
    
End Sub

private sub exportaReferencias(db,ts)
    dim ret
    dim ref
    dim columna
    for each ref in db.relations
        ts.write "ALTER TABLE " & ref.foreigntable & " ADD CONSTRAINT fk_" & ref.foreigntable & "_" & ref.table & " FOREIGN KEY(" 
        ret=""
    
        for each columna in ref.fields
            ret= ret & "," & columna.name
        next
        ts.write mid(ret,2)
        ts.write ") REFERENCES " & ref.table & colsClave(db.tabledefs(ref.table)) 
        if ref.Attributes and dbRelationDeleteCascade then
          ts.write " ON DELETE CASCADE"
        end if
        if ref.Attributes and dbRelationUpdateCascade then
          ts.write " ON UPDATE CASCADE"
        end if
        ts.writeline ";"
        if bExportTrigger then
          call createTrigger(db, ref, ts)
        end if
    next
end sub

function get_trigger_name(trigger_type, ptname, pcname, ftname, fcname)
  get_trigger_name = trigger_type & "_" & ptname & "_" & pcname & "_" & ftname & "_" & fcname
end function

function null_sql(fcname, nullable)
    null_sql = ""
    if not nullable then
        null_sql = "NEW." & fcname & " IS NOT NULL AND "
    end if
end function

sub createTrigger(db, ref, ts)
  dim trigger_name
  dim ptname, pcname, ftname, fcname, col
  Dim str1, str2
  Dim print_string 

    ptname = ref.table
    pcname = ref.Fields(0).Name
    ftname = ref.ForeignTable
    fcname = ref.Fields(0).ForeignName
    '/* table */
    ts.writeline 
    trigger_name = get_trigger_name("fki", ptname, pcname, ftname, fcname)
    print_string = "CREATE TRIGGER " & trigger_name & " BEFORE INSERT ON " & ftname & _
                   " FOR EACH ROW BEGIN" & _
                   " SELECT RAISE(ABORT, 'insert on table " & ftname & " violates foreign key constraint " & trigger_name & "')"
    print_string = print_string & " WHERE " 
    str1="" : str2=""
    for each col in ref.Fields
      str1 = str1 & null_sql(col.ForeignName, db.TableDefs(ftname).Fields(col.ForeignName).Required) 
      str2 = str2 & " AND " & col.Name & "=NEW." & col.ForeignName
    next
    print_string = print_string & str1 & "(SELECT " & pcname & " FROM " & ptname & " WHERE " & Mid(str2, 6) & ") IS NULL;"
    ts.writeline print_string & "END;" 
    '------------------------------------
    if ref.Attributes and dbRelationUpdateCascade then
      trigger_name = get_trigger_name("fkuc", ptname, pcname, ftname, fcname)
      print_string = "CREATE TRIGGER " & trigger_name & " AFTER UPDATE ON " & ptname & _
                     " FOR EACH ROW BEGIN"      
      str1="" : str2=""
      for each col in ref.Fields
        str1 = str1 & "," & col.ForeignName & "=NEW." & col.Name
        str2 = str2 & " AND " & col.ForeignName & "=OLD." & col.Name
      next      
      print_string = print_string & " UPDATE " & ftname & " SET " & Mid(str1, 2) & _
                    "  WHERE " & Mid(str2, 6) & ";"
      ts.writeline print_string & "END;"
    end if
    '------------------------------------
    trigger_name = get_trigger_name("fku", ptname, pcname, ftname, fcname)
    print_string = "CREATE TRIGGER " & trigger_name & " BEFORE UPDATE ON " & ftname & _
                   " FOR EACH ROW BEGIN" & _
                   " SELECT RAISE(ABORT, 'update on table " & ftname & " violates foreign key constraint " & trigger_name & "')"
    print_string = print_string & " WHERE " 
    str1="" : str2=""
    for each col in ref.Fields
      str1 = str1 & null_sql(col.ForeignName, db.TableDefs(ftname).Fields(col.ForeignName).Required) 
      str2 = str2 & " AND " & col.Name & "=NEW." & col.ForeignName
    next
    print_string = print_string & str1 & "(SELECT " & pcname & " FROM " & ptname & " WHERE " & Mid(str2, 6) & ") IS NULL;"
    ts.writeline print_string & "END;"
    '------------------------------------
    ' foreign table 
    if ref.Attributes and dbRelationDeleteCascade then
      'CASCADE
      trigger_name = get_trigger_name("fkdc", ptname, pcname, ftname, fcname)
      print_string = "CREATE TRIGGER " & trigger_name & " BEFORE DELETE ON " & ptname & _
                     " FOR EACH ROW BEGIN" & _
                     " DELETE FROM " & ftname & " WHERE " 
      str2=""
      for each col in ref.Fields     
        str2 = str2 & " AND " & ftname & "." & col.ForeignName & "=OLD." & col.Name
      next
      print_string = print_string & Mid(str2, 6) & ";"
      ts.writeline print_string & "END;"
    else
      trigger_name = get_trigger_name("fkd", ptname, pcname, ftname, fcname)
      print_string = "CREATE TRIGGER " & trigger_name & " BEFORE DELETE ON " & ptname & _
                     " FOR EACH ROW BEGIN" & _
                     " SELECT RAISE(ABORT, 'delete on table " & ptname & " violates foreign key constraint " & trigger_name & "')" & _
                     " WHERE " 
      str2=""
      for each col in ref.Fields     
        str2 = str2 & " AND " & col.ForeignName & "=OLD." & col.Name
      next
      print_string = print_string & "(SELECT " & fcname & " FROM " & ftname & " WHERE " & Mid(str2, 6) & ") IS NOT NULL;"
      ts.writeline print_string & "END;"
    end if
end sub

private function colsClave(tabla)
On Error Resume Next
    dim clave
    dim print_string
    dim col
    set clave = tabla.indexes("PrimaryKey")
    print_string=""
    for each col in clave.fields
        print_string =  print_string & "," & col.name 
    next 
    'WScript.Echo print_string 
    
    colsClave="(" & mid(print_string ,2) & ")"

end function

private sub exportIndexes(tabla, ts)
On Error Resume Next
    dim clave
    dim print_string
    dim col
    dim cols
    dim bUniq
    dim idx
    for each idx in tabla.indexes
      print_string=""
      for each col in idx.fields
        print_string =  print_string & "," & col.name 
      next 
      cols="(" & mid(print_string ,2) & ")"
      if not idx.Primary then 'and not idx.Foreign then 
        if idx.Unique then 
          print_string =  "CREATE UNIQUE INDEX "
        else
          print_string =  "CREATE INDEX "
        end if
        ts.writeline print_string & "idx_" & tabla.name & "_" & idx.name & " ON " & tabla.name & cols & ";"
      end if
    next 
end sub

Private Sub exportDatabase(ByVal database_path , ByVal username , ByVal password , _
                           ByVal outfile )
    Dim fso,ts
    Dim print_string
    Dim db 
    Dim tabla, columna
    Dim wrkJet 
    Dim table_name 
    Dim table_sql_name 
    Dim field_name 
    Dim field_sql_name 
    Dim field_type 
    dim dbEng
    dim col
    dim ref
    dim bAutoInc
    
    ' Create Microsoft Jet Workspace object.
    set dbEng=createobject("DAO.DBEngine.36")
    Set wrkJet = dbEng.CreateWorkspace("", username, password, dbUseJet)

    Set db = wrkJet.OpenDatabase(database_path, , True)
    
    If (outfile = "") Then outfile = "sqlite_db_out.sql"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(outfile, True)
    
    'ts.writeline "SET DEFINE OFF;" 'Así no fastidian los caracteres &
    For each tabla in db.TableDefs
    		if LCASE(LEFT(tabla.name,4)) <> "msys" Then
	      table_name = tabla.name
        table_sql_name = sql_name(table_name)
        If Left(table_name, 4) <> "MSys" Then
            WScript.Echo tabla.name
            bAutoInc = False
            ts.writeline "BEGIN TRANSACTION;"
            ts.writeline "DROP TABLE IF EXISTS " & table_sql_name & ";"
            print_string = "CREATE TABLE " & table_sql_name & " ( "
            For each columna in tabla.Fields
                field_name = columna.name
                field_sql_name = sql_name(field_name)
                field_type = getSQLiteFieldType(columna, False)
                if field_type <> "BINARY" then
                    print_string = print_string & field_sql_name & " "
                    print_string = print_string & field_type 
                    if columna.Attributes and dbAutoIncrField then
                        print_string = print_string & " PRIMARY KEY"
                        bAutoInc = True
                    end if
                    if columna.required then
                        print_string = print_string & " NOT NULL"
                    end if
                    if columna.defaultValue <> "" then
                      if columna.DefaultValue = "Now()" then
                        print_string = print_string & " DEFAULT CURRENT_TIMESTAMP"
                      else
                        print_string = print_string & " DEFAULT " & columna.DefaultValue
                      end if
                    end if
                    print_string = print_string & ", "
                end if
            Next 
            ts.write Left(print_string, Len(print_string) - 2)
            if not bAutoInc and colsClave(tabla) <> "()" then
              ts.writeline ","
              ts.write " primary key "
              ts.write colsClave(tabla)
            end if
            ts.writeline ");"
            
            For each columna in tabla.Fields
                field_type = getSQLiteFieldType(columna, False)
                If InStr(field_type, "-- error") = 1 Then
                    ts.writeline  field_type 
                End If
            Next 
            
            exportIndexes tabla, ts

            if bExportInsert then
                Call exportDatabaseTable(db, tabla, ts)
            end if
            ts.writeline "COMMIT TRANSACTION;"
        End If
    		End if
    Next 
    
    exportaReferencias db,ts

    'ts.writeline ".Q" 'Si no, se queda
    ts.close
    
    db.Close
    wrkJet.Close
    Exit Sub
    
DB_Error:
    MsgBox "Could not open database " & database_path & ". Check the path and permissions."
        
End Sub

Public Sub Main(mdb_db)
    
    Dim outfile
    'keywords i took from sqlite/tokenizer.c
    sql_keywords1 = Array( _
        "ABORT", "AFTER", "ALL", "AND", "AS", "ASC", "ATTACH", _
        "BEFORE", "BEGIN", "BETWEEN", "BY", "CASCADE", "CASE", _
        "CHECK", "CLUSTER", "COLLATE", "COMMIT", "CONFLICT", _
        "CONSTRAINT", "COPY", "CREATE", "CROSS", "DATABASE", _
        "DEFAULT", "DEFERRED", "DEFERRABLE", "DELETE", _
        "DELIMITERS", "DESC", "DETACH", "DISTINCT", "DROP", _
        "END", "EACH", "ELSE", "EXCEPT", "EXPLAIN", "FAIL", _
        "FOR", "FOREIGN", "FROM", "FULL", "GLOB", "GROUP" _
        )
    sql_keywords2 = Array( _
        "HAVING", "IGNORE", "IMMEDIATE", "IN", "INDEX", _
        "INITIALLY", "INNER", "INSERT", "INSTEAD", "INTERSECT", _
        "INTO", "IS", "ISNULL", "JOIN", "KEY", "LEFT", "LIKE", _
        "LIMIT", "MATCH", "NATURAL", "NOT", "NOTNULL", "NULL", _
        "OF", "OFFSET", "ON", "OR", "ORDER", "OUTER", "PRAGMA", _
        "PRIMARY", "RAISE", "REFERENCES", "REPLACE", "RESTRICT", _
        "RIGHT", "ROLLBACK", "ROW", _
        "SELECT", "SET", "STATEMENT", "TABLE", "TEMP", _
        "TEMPORARY", "THEN", "TRANSACTION", "TRIGGER", _
        "UNION", "UNIQUE", "UPDATE", "USING", "VACUUM", _
        "VALUES", "VIEW", "WHEN", "WHERE" _
        )

    
    If Right(mdb_db, 4) <> ".mdb" Then
        outfile = mdb_db & ".sql"
    Else
        outfile = Replace(mdb_db, ".mdb", ".sql")
    End If
    
    Call exportDatabase(mdb_db, "admin", "", outfile)

End Sub

Dim arg
WScript.Echo "MDBtoSQLite.vbs (c) 2008 <lokkju@lokkju.com>,others"
WScript.Echo " This program comes with ABSOLUTELY NO WARRANTY;"
WScript.Echo " This is free software, and you are welcome to redistribute it"
WScript.Echo "  under certain conditions; view this file as text for details."
WScript.Echo ""
If WScript.Arguments.Count = 0 Then
   WScript.Echo  "Usage:"
   WScript.Echo "    cscript.exe MDBtoSQLite.vbs <mdb file> [second mdb file] [...]"
   WScript.Echo "    Example: cscript MDBtoSQLite.vbs table.mdb"
   WScript.Echo "    Adding more arguments will process each given file"
Else
   For each arg in WScript.Arguments  
      main arg
      Exit For
   Next
End If
