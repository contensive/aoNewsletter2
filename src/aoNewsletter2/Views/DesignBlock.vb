
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
    Public Class TileClass
        Inherits AddonBaseClass
        '
        '====================================================================================================
        '
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Const designBlockName As String = "Tile Design Block"
            Try
                '
                ' -- read instanceId, guid created uniquely for this instance of the addon on a page
                Dim result = String.Empty
                Dim settingsGuid = DesignBlockController.getSettingsGuid(CP, designBlockName, result)
                If (String.IsNullOrEmpty(settingsGuid)) Then Return result
                '
                ' -- locate or create a data record for this guid
                Dim settings = DbTileModel.createOrAddSettings(CP, settingsGuid)
                If (settings Is Nothing) Then Throw New ApplicationException("Could not create the design block settings record.")
                '
                ' -- translate the Db model to a view model and mustache it into the layout
                Dim viewModel = TileViewModel.create(CP, settings)
                If (viewModel Is Nothing) Then Throw New ApplicationException("Could not create design block view model.")
                result = Nustache.Core.Render.StringToString(My.Resources.TileLayout, viewModel)
                '
                ' -- if editing enabled, add the link and wrapperwrapper
                Return GenericController.addEditWrapper(CP, result, settings.id, settings.name, DbTileModel.contentName)
            Catch ex As Exception
                CP.Site.ErrorReport(ex)
                Return "<!-- " & designBlockName & ", Unexpected Exception -->"
            End Try
        End Function
    End Class
End Namespace