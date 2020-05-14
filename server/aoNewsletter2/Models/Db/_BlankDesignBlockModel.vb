Imports System
Imports Contensive.Addons.Newsletter.Models.Db
Imports Contensive.BaseClasses
Imports Contensive.Models.Db

Namespace Models.Db
    Public Class BlankDesignBlockModel
        Inherits DesignBlockBaseModel

        Public Shared ReadOnly Property tableMetadata As DbBaseTableMetadataModel = New DbBaseTableMetadataModel("blank", "blank", "default", False)
        Public Property imageFilename As String
        Public Property headline As String
        Public Property embed As String
        Public Property description As String
        Public Property buttonText As String
        Public Property buttonUrl As String
        Public Property imageAspectRatioId As Integer
        Public Property btnStyleSelector As String

        Public Shared Function createOrAddSettings(ByVal cp As CPBaseClass, ByVal settingsGuid As String) As BlankDesignBlockModel
            Dim result As BlankDesignBlockModel = create(Of BlankDesignBlockModel)(cp, settingsGuid)

            If (result Is Nothing) Then
                result = DesignBlockBaseModel.addDefault(Of BlankDesignBlockModel)(cp)
                result.name = tableMetadata.contentName & " " & result.id
                result.ccguid = settingsGuid
                result.themeStyleId = 0
                result.padTop = False
                result.padBottom = False
                result.padRight = False
                result.padLeft = False
                result.imageFilename = String.Empty
                result.imageAspectRatioId = 3
                result.headline = "Lorem Ipsum Dolor"
                result.description = "<p>Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.</p>"
                result.embed = String.Empty
                result.buttonUrl = String.Empty
                result.buttonText = String.Empty
                result.save(cp)
                cp.Content.LatestContentModifiedDate.Track(result.modifiedDate)
            End If

            Return result
        End Function
    End Class
End Namespace

