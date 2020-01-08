
Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace newsletter2
    Public Class genericController
        '
        '====================================================================================================
        ''' <summary>
        ''' return date.minValue if date is before 1/1/1900
        ''' </summary>
        ''' <param name="sourceDate"></param>
        ''' <returns></returns>
        Friend Shared Function encodeMinDate(sourceDate As Date) As Date
            If (sourceDate < #1/1/1900#) Then
                Return Date.MinValue
            End If
            Return sourceDate
        End Function
    End Class
End Namespace
