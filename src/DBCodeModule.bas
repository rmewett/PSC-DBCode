Attribute VB_Name = "DBCodeModule"
Option Explicit

Public mDb As DAO.Database
Public Function GetSQLType(DataType As Long) As String
    Select Case DataType
    Case dbBinary
        GetSQLType = "BINARY"
    Case dbBoolean
        GetSQLType = "BIT"
    Case dbByte
        GetSQLType = "BYTE"
    Case dbCurrency
        GetSQLType = "CURRENCY"
    Case dbDate
        GetSQLType = "DATETIME"
    Case dbDouble
        GetSQLType = "DOUBLE"
    Case dbInteger
        GetSQLType = "SHORT"
    Case dbLong
        GetSQLType = "LONG"
    Case dbMemo
        GetSQLType = "LONGTEXT"
    Case dbSingle
        GetSQLType = "SINGLE"
    Case dbText
        GetSQLType = "TEXT"
    Case dbTime
        GetSQLType = "DATETIME"
    End Select
End Function


Public Function GetDBConstant(lValue As Long) As String
    Select Case lValue
    Case dbBigInt
        GetDBConstant = "dbBigInt"
    Case dbBinary
        GetDBConstant = "dbBinary"
    Case dbBoolean
        GetDBConstant = "dbBoolean"
    Case dbByte
        GetDBConstant = "dbByte"
    Case dbChar
        GetDBConstant = "dbByte"
    Case dbCurrency
        GetDBConstant = "dbCurrency"
    Case dbDate
        GetDBConstant = "dbDate"
    Case dbDecimal
        GetDBConstant = "dbDecimal"
    Case dbDouble
        GetDBConstant = "dbDouble"
    Case dbFloat
        GetDBConstant = "dbFloat"
    Case dbGUID
        GetDBConstant = "dbGUID"
    Case dbInteger
        GetDBConstant = "dbInteger"
    Case dbLong
        GetDBConstant = "dbLong"
    Case dbLongBinary
        GetDBConstant = "dbLongBinary"
    Case dbMemo
        GetDBConstant = "dbMemo"
    Case dbNumeric
        GetDBConstant = "dbNumeric"
    Case dbSingle
        GetDBConstant = "dbSingle"
    Case dbText
        GetDBConstant = "dbText"
    Case dbTime
        GetDBConstant = "dbTime"
    Case dbTimeStamp
        GetDBConstant = "dbTimeStamp"
    Case dbVarBinary
        GetDBConstant = "dbVarBinary"
    Case Else
        GetDBConstant = CStr(lValue)
    End Select
End Function



