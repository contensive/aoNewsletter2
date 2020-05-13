Imports System
Imports Contensive.Addons.Newsletter.Models.Db
Imports Contensive.BaseClasses
Imports Contensive.Models.Db

Namespace Models.Db
    Public Class NewsletterModel
        Inherits DesignBlockBaseModel

        Public Shared ReadOnly Property tableMetadata As DbBaseTableMetadataModel = New DbBaseTableMetadataModel("Newsletters", "Newsletters", "default", False)
        '
        Public Property templateId As Integer
        Public Property stylesFileName As DbBaseModel.FieldTypeCSSFile
        Public Property emailTemplateId As Integer
        Public Property mastheadFilename As String
        Public Property footerFilename As String
        Public Property blockArchiveSearchForm As Boolean
        Public Property archiveIssuesToDisplay As Integer
        Public Property searchResultsPerPage As Integer

        Public Shared Function createOrAddSettings(ByVal cp As CPBaseClass, ByVal settingsGuid As String) As NewsletterModel
            Dim result As NewsletterModel = create(Of NewsletterModel)(cp, settingsGuid)

            If (result Is Nothing) Then
                result = NewsletterModel.addDefault(Of NewsletterModel)(cp)
                result.name = NewsletterModel.tableMetadata.contentName & " " & result.id
                result.ccguid = settingsGuid
                result.themeStyleId = 0
                result.padTop = False
                result.padBottom = False
                result.padRight = False
                result.padLeft = False

                result.save(cp)
                cp.Content.LatestContentModifiedDate.Track(result.modifiedDate)
            End If

            Return result
        End Function
    End Class
End Namespace

