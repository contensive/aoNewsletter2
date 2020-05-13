
Imports Contensive.Addons.Newsletter.Controllers
Imports Contensive.BaseClasses


Namespace Models.View
    Public Class NewsletterViewModel
        Inherits DesignBlockViewBaseModel
        '
        Public Property legacyNewsletter As String
        '
        '====================================================================================================
        ''' <summary>
        ''' Populate the view model from the entity model
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <param name="settings"></param>
        ''' <returns></returns>
        Public Overloads Shared Function create(cp As CPBaseClass, settings As Models.Db.NewsletterModel, legacyNewsletter As String) As NewsletterViewModel
            Try
                '
                ' -- base fields
                Dim result = DesignBlockViewBaseModel.create(Of NewsletterViewModel)(cp, settings)
                '
                ' -- custom
                result.legacyNewsletter = legacyNewsletter
                '
                Return result
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
                Return Nothing
            End Try
        End Function
    End Class

End Namespace