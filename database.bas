Attribute VB_Name = "database"
Option Explicit

Public Declare Sub sqlite3_open Lib "sqlite.dll" (ByVal FileName As String, ByRef handle As Long)
Public Declare Sub sqlite3_close Lib "sqlite.dll" (ByVal DB_Handle As Long)
Public Declare Function sqlite3_last_insert_rowid Lib "sqlite.dll" (ByVal DB_Handle As Long) As Long
Public Declare Function sqlite3_changes Lib "sqlite.dll" (ByVal DB_Handle As Long) As Long
Public Declare Function sqlite_get_table Lib "sqlite.dll" (ByVal DB_Handle As Long, ByVal SQLString As String, ByRef ErrStr As String) As Variant()
Public Declare Function sqlite_libversion Lib "sqlite.dll" () As String
Public Declare Function number_of_rows_from_last_call Lib "sqlite.dll" () As Long

Public DBz As Long
Public DBFile As String
Public minfo As String ' sql error akan store kat sini
Public row As Variant
Public query As String ' public variable untuk sql query
Public numrows As Long
Public i As Long

'connect ke database
Public Function connectDb()
DBFile = App.Path & "\database.db"
sqlite3_open DBFile, DBz

End Function

'tutup database
Public Function closeDB()

    sqlite3_close (DBz)

End Function


Public Function getData()
i = 1 'initialize counter untuk show data

query = "SELECT * FROM users"

row = sqlite_get_table(DBz, query, minfo) ' query database
numrows = number_of_rows_from_last_call ' bilangan rows data yang di select
End Function


