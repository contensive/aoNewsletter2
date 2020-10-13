
Imports System
Imports Contensive.Addons.Newsletter.Controllers
Imports Contensive.Addons.Newsletter.Models.Db
Imports Contensive.BaseClasses
Imports Contensive.Models.Db


Namespace Models.View
    Public Class DesignBlockViewBaseModel
        Public Property styleBackgroundImage As String
        Public Property styleheight As String
        Public Property contentContainerClass As String
        Public Property outerContainerClass As String
        '
        '====================================================================================================
        ''' <summary>
        ''' Populate the view model from the entity model
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <param name="settings"></param>
        ''' <returns></returns>
        Public Shared Function create(Of T As DesignBlockViewBaseModel)(cp As CPBaseClass, settings As Models.Db.DesignBlockBaseModel) As T
            Dim result As T = Nothing
            Try
                Dim instanceType As Type = GetType(T)
                result = DirectCast(Activator.CreateInstance(instanceType), T)
                '
                ' -- base fields
                result.styleheight = encodeStyleHeight(settings.styleheight)
                result.styleBackgroundImage = "" _
                    & encodeStyleBackgroundImage(cp, settings.backgroundImageFilename) _
                    & ""
                result.outerContainerClass = "" _
                    & If(settings.themeStyleId.Equals(0), String.Empty, " " & cp.Content.GetRecordName(DesignBlockThemeModel.tableMetadata.contentName, settings.themeStyleId)) _
                    & ""
                result.contentContainerClass = "" _
                    & If(settings.asFullBleed, " container", String.Empty) _
                    & If(settings.padTop, " pt-5", " pt-0") _
                    & If(settings.padRight, " pr-4", " pr-0") _
                    & If(settings.padBottom, " pb-5", " pb-0") _
                    & If(settings.padLeft, " pl-4", " pl-0") _
                    & ""
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '====================================================================================================
        ''' <summary>
        ''' convert string into a style "height: {styleHeight};", if value is numeric it adds "px"
        ''' </summary>
        ''' <param name="styleheight"></param>
        ''' <returns></returns>
        Public Shared Function encodeStyleHeight(styleheight As String) As String
            Return If(String.IsNullOrWhiteSpace(styleheight), String.Empty, "overflow:hidden;height:" & styleheight & If(IsNumeric(styleheight), "px", String.Empty) & ";")
        End Function
        '
        '====================================================================================================
        ''' <summary>
        ''' convert string into a style "background-image: url(backgroundImage)
        ''' </summary>
        ''' <param name="backgroundImage"></param>
        ''' <returns></returns>
        Public Shared Function encodeStyleBackgroundImage(cp As CPBaseClass, backgroundImage As String) As String
            Return If(String.IsNullOrWhiteSpace(backgroundImage), String.Empty, "background-image: url('" & cp.Http.CdnFilePathPrefix & backgroundImage & "');")
        End Function
    End Class
End Namespace
