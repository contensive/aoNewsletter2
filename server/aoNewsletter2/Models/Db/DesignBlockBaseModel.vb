
Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace Models.Db
    '
    ''' <summary>
    ''' This model provides the common fields for all Design Blocks.
    ''' </summary>
    Public Class DesignBlockBaseModel
        Inherits Global.Contensive.Models.Db.DbBaseModel
        '
        '====================================================================================================
        ' -- instance properties
        '
        Public Property backgroundImageFilename As String
        Public Property fontStyleId As Integer
        Public Property themeStyleId As Integer
        Public Property padTop As Boolean
        Public Property padBottom As Boolean
        Public Property padRight As Boolean
        Public Property padLeft As Boolean
        Public Property styleheight As String
        Public Property asFullBleed As Boolean
    End Class
End Namespace
