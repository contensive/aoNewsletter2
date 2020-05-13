
Imports System
Imports Contensive.Addons.Newsletter.Controllers
Imports Contensive.Addons.Newsletter.Models.Db
Imports Contensive.BaseClasses
Imports Contensive.Models.Db

Namespace Contensive.Addons.SampleCollection
    Namespace Models.View
        Public Class DesignBlockViewBaseModel
            Public Property styleBackgroundImage As String
            Public Property styleheight As String
            Public Property contentContainerClass As String
            Public Property outerContainerClass As String

            Public Shared Function create(Of T As DesignBlockViewBaseModel)(ByVal cp As CPBaseClass, ByVal settings As DesignBlockBaseModel) As T
                Dim result As T = Nothing

                Try
                    Dim instanceType As Type = GetType(T)
                    result = CType(Activator.CreateInstance(instanceType), T)
                    result.styleheight = encodeStyleHeight(settings.styleheight)
                    result.styleBackgroundImage = "" & encodeStyleBackgroundImage(cp, settings.backgroundImageFilename) & ""
                    result.outerContainerClass = "" & (If(settings.fontStyleId.Equals(0), String.Empty, " ")) & DbBaseModel.getRecordName(Of DesignBlockFontModel)(cp, settings.fontStyleId) & (If(settings.themeStyleId.Equals(0), String.Empty, " ")) & ""
                    result.contentContainerClass = "" & (If(settings.asFullBleed, " container", String.Empty)) & (If(settings.padTop, " pt-5", " pt-0")) & (If(settings.padRight, " pr-4", " pr-0")) & (If(settings.padBottom, " pb-5", " pb-0")) & (If(settings.padLeft, " pl-4", " pl-0")) & ""
                Catch ex As Exception
                    cp.Site.ErrorReport(ex)
                End Try

                Return result
            End Function

            Public Shared Function encodeStyleHeight(ByVal styleheight As String) As String
                Return If(String.IsNullOrWhiteSpace(styleheight), String.Empty, "overflow:hidden;height:" & styleheight & (If(GenericController.isNumeric(styleheight), "px", String.Empty)) & ";")
            End Function

            Public Shared Function encodeStyleBackgroundImage(ByVal cp As CPBaseClass, ByVal backgroundImage As String) As String
                Return If(String.IsNullOrWhiteSpace(backgroundImage), String.Empty, "background-image: url('" & cp.Site.FilePath & backgroundImage & "');")
            End Function
        End Class
    End Namespace
End Namespace
