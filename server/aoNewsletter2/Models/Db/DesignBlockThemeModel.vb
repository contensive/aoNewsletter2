
Imports Contensive.Models.Db

Namespace Models
    Public Class DesignBlockThemeModel
        Inherits DbBaseModel

        Public Shared ReadOnly Property tableMetadata As DbBaseTableMetadataModel = New DbBaseTableMetadataModel("Design Block Themes", "dbThemes", "default", False)

    End Class
End Namespace