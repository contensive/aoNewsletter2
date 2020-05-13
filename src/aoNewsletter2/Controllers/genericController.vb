
Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Contensive.BaseClasses

Namespace Controllers
    Public Class GenericController
        '
        '====================================================================================================
        ''' <summary>
        ''' return date.minValue if date is before 1/1/1900
        ''' </summary>
        ''' <param name="sourceDate"></param>
        ''' <returns></returns>
        Public Shared Function encodeMinDate(sourceDate As Date) As Date
            If (sourceDate < #1/1/1900#) Then
                Return Date.MinValue
            End If
            Return sourceDate
        End Function

        Public Shared Function isNumeric(ByVal value As String) As Boolean
            Return value.All(AddressOf Char.IsNumber)
        End Function

        Public Shared Function isDateEmpty(ByVal srcDate As DateTime) As Boolean
            Return (srcDate < New DateTime(1900, 1, 1))
        End Function

        Public Shared Function getShortDateString(ByVal srcDate As DateTime) As String
            If Not isDateEmpty(srcDate) Then
                Return encodeMinDate(srcDate).ToShortDateString()
            End If

            Return String.Empty
        End Function

        Public Shared Function getSortOrderFromInteger(ByVal id As Integer) As String
            Return id.ToString().PadLeft(7, "0"c)
        End Function

        Public Shared Function getDateForHtmlInput(ByVal source As DateTime) As String
            If isDateEmpty(source) Then
                Return ""
            Else
                Return source.Year & "-" & source.Month.ToString().PadLeft(2, "0"c) & "-" & source.Day.ToString().PadLeft(2, "0"c)
            End If
        End Function

        Public Shared Function verifyProtocol(ByVal url As String) As String
            If (String.IsNullOrWhiteSpace(url)) Then Return String.Empty
            If (url.Substring(0, 1) = "/") Then Return url
            If (Not url.IndexOf("://").Equals(-1)) Then Return url
            Return "http://" & url
        End Function

    End Class
End Namespace