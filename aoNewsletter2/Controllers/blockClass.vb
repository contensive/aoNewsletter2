
Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace newsletter2
    ''' <summary>
    ''' A replacement for cp.block using htmlAgility
    ''' </summary>
    Public Class blockClass
        '
        Dim layout2 As HtmlAgilityPack.HtmlDocument
        '
        Public Sub load(html As String)
            layout2 = New HtmlAgilityPack.HtmlDocument
            layout2.LoadHtml(html)
        End Sub
        '
        Public Sub prepend(html As String)
            layout2.Load(html & getHtml())
        End Sub
        '
        Public Sub append(html As String)
            layout2.Load(getHtml() & html)
        End Sub
        '
        Public Function getClassInner(findClass As String) As String
            Dim node As HtmlAgilityPack.HtmlNode = layout2.DocumentNode.SelectSingleNode("//*[@class='" & findClass & "']")
            If Not (node Is Nothing) Then
                Return node.InnerHtml()
            End If
            Return String.Empty
        End Function
        '
        Public Function getClassOuter(findClass As String) As String
            Dim node As HtmlAgilityPack.HtmlNode = layout2.DocumentNode.SelectSingleNode("//*[@class='" & findClass & "']")
            If Not (node Is Nothing) Then
                Return node.OuterHtml()
            End If
            Return String.Empty
        End Function
        '
        Public Sub setClassInner(findClass As String, replacement As String)
            Dim nodes As HtmlAgilityPack.HtmlNodeCollection = layout2.DocumentNode.SelectNodes("//*[@class='" & findClass & "']")
            If Not (nodes Is Nothing) Then
                For Each node As HtmlAgilityPack.HtmlNode In nodes
                    node.InnerHtml = replacement
                Next
            End If
        End Sub
        '
        Public Sub setClassOuter(findClass As String, replacement As String)
            Dim nodes As HtmlAgilityPack.HtmlNodeCollection = layout2.DocumentNode.SelectNodes("//*[@class='" & findClass & "']")
            If Not (nodes Is Nothing) Then
                For Each node As HtmlAgilityPack.HtmlNode In nodes
                    Dim newNode As HtmlAgilityPack.HtmlNode = HtmlAgilityPack.HtmlNode.CreateNode(replacement)
                    node.ParentNode.ReplaceChild(newNode, node)
                Next
            End If
        End Sub
        '
        Public Function getHtml() As String
            If Not layout2 Is Nothing Then
                If Not (layout2.DocumentNode Is Nothing) Then
                    Return layout2.DocumentNode.OuterHtml
                End If
            End If
            Return String.Empty
        End Function
    End Class
End Namespace

