
Imports Contensive.Models.Db

Namespace Contensive.Addons.SampleCollection
    Namespace Models
        Public Class DesignBlockFontModel
            Inherits DbBaseModel

            Public Shared ReadOnly Property tableMetadata As DbBaseTableMetadataModel = New DbBaseTableMetadataModel("Design Block Fonts", "dbfonts", "default", False)

        End Class
    End Namespace
End Namespace
