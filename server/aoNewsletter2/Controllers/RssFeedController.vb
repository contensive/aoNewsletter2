Imports System.Linq.Expressions
Imports Contensive.BaseClasses

Public Class RssFeedController
    '
    Public Shared Sub updateRSSFeed(cp As CPBaseClass)
        '
        cp.Db.ExecuteNonQuery($"upadte ccaggregatefunctions set processRunOnce=1 where name='RSS Feed Process'")
        '
    End Sub
End Class
