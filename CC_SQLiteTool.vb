Imports System.Data.SQLite
Public Class CC_SQLiteTool
    Private conn As New SQLiteConnection
    Private sqlcmd As New SQLiteCommand
    Private sqlreader As SQLiteDataReader
    Private adapter As SQLiteDataAdapter
    Private OK As Integer = 1
    Private NG As Integer = -1
    Private Simplecommand As Simple_SQLiteCommand = New Simple_SQLiteCommand
    ''' <summary>
    ''' 建立数据库文件
    ''' </summary>
    ''' <param name="DBpath">数据库文件路径，通常以".db"结尾</param>
    ''' <returns></returns>
    Public Function CreatDB(DBpath As String) As Integer
        Try
            conn.ConnectionString = "Data Source=" & DBpath
            conn.Open()
            conn.Close()
            Return OK
        Catch ex As Exception
            Return NG
        End Try
    End Function

    ''' <summary>
    ''' 创建数据库中的表
    ''' 创建表中的列的字符串规则相对比较繁琐，具体规则如下：
    ''' 1、每一列分为“列名”和“数据类型”，列名和数据类型中间用空格隔开；
    ''' 2、列和列之间使用“,”（英文逗号）分隔，逻辑上N列应该有N-1个逗号
    ''' 3、列名可以是中文；
    ''' 4、类型共有：“NULL”、“INTEGER”、“REAL”、“TEXT”、“BLOB”，这是总的类型；
    ''' 5、Null：此为空类型，无扩展
    ''' 6、INTEGER：此为整型，通常可以扩展为：INT/INTEGER/TINYINT/SMALLINT/BIGINT/UNSIGED BIG INT/INT2/INT8
    ''' 7、TEXT：此为字符串，通常可扩展为：CHAR(10)/CHARACTER(20)/VARCHAR(255)/VARYING CHARACTER(255)/NCHAR(25)/TEXT/CLOB/NVARCHAR(100)
    ''' 8、REAL：此为浮点型，通常可扩展为：REAL/DOUBLE/FLOAT/DOUBLE PRECISION
    ''' 9、NUMERIC：此可能为BLOB块，通常扩展为：NUMERIC/DECIMAL(10.5)/BOOLEAN/DATE/DATETIME
    ''' </summary>
    ''' <param name="DBpath">数据库文件路径</param>
    ''' <param name="TableName">需要创建的表的表名</param>
    ''' <param name="columnString">创建所有的列的字符串，此规则比较繁琐，请仔细确认</param>
    ''' <returns></returns>
    Public Function CreatTable(DBpath As String, TableName As String, columnString As String) As Integer
        Try
            conn.ConnectionString = "Data Source=" & DBpath
            conn.Open()
            sqlcmd.Connection = conn
            sqlcmd.CommandText = String.Format("CREATE TABLE IF NOT EXISTS {0} ({1})", TableName, columnString)
            sqlcmd.ExecuteNonQuery()
            conn.Close()
            Return OK
        Catch ex As Exception
            Return NG
        End Try
    End Function

    Public Function ConnectDatabase(DBpath As String) As Integer
        Try
            conn.ConnectionString = "Data Source=" & DBpath
            conn.Open()
            sqlcmd.Connection = conn
            Return OK
        Catch ex As Exception
            Return NG
        End Try
    End Function

    Public Function Close() As Integer
        Try
            conn.Close()
            Return OK
        Catch ex As Exception
            Return NG
        End Try
    End Function

    ''' <summary>
    ''' 在数据库中增加一行
    ''' </summary>
    ''' <param name="TableName">表名</param>
    ''' <param name="Values">增加的数据内容，以字符串列表传入</param>
    ''' <returns></returns>
    Public Function InsertRow(TableName As String, Values As List(Of String)) As Integer
        Try
            Dim myvaluestring As String = ""
            For Each myValue In Values
                myvaluestring += "'" + myValue + "'" + ","
            Next
            myvaluestring = myvaluestring.Remove(myvaluestring.Length - 1, 1)
            sqlcmd.CommandText = String.Format("INSERT INTO {0} VALUES ({1})", TableName, myvaluestring)
            sqlcmd.ExecuteNonQuery()
            Return OK
        Catch ex As Exception
            Return NG
        End Try
    End Function

    ''' <summary>
    ''' 删除表中的前几行
    ''' </summary>
    ''' <param name="TableName">表命</param>
    ''' <param name="LineNum">行数</param>
    ''' <returns></returns>
    Public Function DeleteFirstFewLines(TableName As String, LineNum As Integer) As Integer
        Try
            sqlcmd.CommandText = Simplecommand.DeleteFirstFewLines(TableName, LineNum)
            sqlcmd.ExecuteNonQuery()
            Return OK
        Catch ex As Exception
            Return NG
        End Try
    End Function

    ''' <summary>
    ''' 获取表中前几行
    ''' </summary>
    ''' <param name="TableName">表命</param>
    ''' <param name="LineNum">行数</param>
    ''' <returns></returns>
    Public Function GetFirstFewLines(TableName As String, LineNum As Integer) As DataTable
        Dim dt As DataTable = New DataTable
        Try
            sqlcmd.CommandText = Simplecommand.GetFirstFewLines(TableName, LineNum)
            adapter = New SQLiteDataAdapter(sqlcmd)
            adapter.AcceptChangesDuringFill = False
            adapter.Fill(dt)
        Catch ex As Exception

        End Try
        Return dt
    End Function

    Public Function GetTableLineNum(TableName As String) As Long
        sqlcmd.CommandText = String.Format("select count(*)  from {0} ", TableName)
        Return Convert.ToInt64(sqlcmd.ExecuteScalar())
    End Function

End Class
Public Class Simple_SQLiteCommand
    ''' <summary>
    ''' 获取表中前几行
    ''' </summary>
    ''' <param name="TableName">表命</param>
    ''' <param name="LineNum">行数</param>
    ''' <returns></returns>
    Public Function GetFirstFewLines(TableName As String, LineNum As Integer) As String
        Return String.Format("SELECT * FROM {0} where rowid in(select rowid from {0} limit {1} offset 0)", TableName, LineNum.ToString())
    End Function

    ''' <summary>
    ''' 删除表中的前几行
    ''' </summary>
    ''' <param name="TableName">表命</param>
    ''' <param name="LineNum">行数</param>
    ''' <returns></returns>
    Public Function DeleteFirstFewLines(TableName As String, LineNum As Integer) As String
        Return String.Format("delete from {0} where rowid in(select rowid from {0} limit {1} offset 0)", TableName, LineNum.ToString())
    End Function
End Class
