
Imports Contensive.Addons.Newsletter.Models.Db
Imports Contensive.Addons.Newsletter.Controllers
Imports Contensive.BaseClasses
Imports Contensive.Addons.Newsletter.Models.View

Namespace Views
    '
    '====================================================================================================
    ''' <summary>
    ''' Design block with a centered headline, image, paragraph text and a button.
    ''' </summary>
    Public Class NewsletterWidget
        Inherits AddonBaseClass
        '
        '====================================================================================================
        '
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Const designBlockName As String = "Newsletter Widget"
            Try
                '
                ' -- read instanceId, guid created uniquely for this instance of the addon on a page
                Dim result = String.Empty
                Dim settingsGuid = DesignBlockController.getSettingsGuid(CP, designBlockName, result)
                If (String.IsNullOrEmpty(settingsGuid)) Then Return result
                '
                ' -- locate or create a data record for this guid
                Dim newsletterId As Integer = NewsletterController.getNewsletterId(CP, settingsGuid)
                Dim currentIssueID As Integer = NewsletterController.GetCurrentIssueID(CP, newsletterId)
                Dim settings = NewsletterModel.createOrAddSettings(CP, settingsGuid)
                If (settings Is Nothing) Then Throw New ApplicationException("Could not create the design block settings record.")
                '
                ' -- create legacy newsletter
                '
                Dim legacyNewsletter As String = (New NewsletterClass()).getLegacyNewsletter(CP, newsletterId, currentIssueID)
                '
                ' -- translate the Db model to a view model and mustache it into the layout
                Dim viewModel = NewsletterViewModel.create(CP, settings, legacyNewsletter)
                If (viewModel Is Nothing) Then Throw New ApplicationException("Could not create design block view model.")
                result = Nustache.Core.Render.StringToString(My.Resources.NewsletterLayout, viewModel)
                '
                ' -- if editing enabled, add the link and wrapperwrapper
                Return CP.Content.GetEditWrapper(result, NewsletterModel.tableMetadata.contentName, settings.id)
            Catch ex As Exception
                CP.Site.ErrorReport(ex)
                Return "<!-- " & designBlockName & ", Unexpected Exception -->"
            End Try
        End Function
    End Class
End Namespace