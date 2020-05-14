Imports System
Imports Contensive.Addons.Newsletter.Models.Db
Imports Contensive.BaseClasses
Imports Contensive.Models.Db

Namespace Models.Db
    Public Class NewsletterAdBannerLayoutModel
        Inherits Global.Contensive.Models.Db.DbBaseModel

        Public Shared ReadOnly Property tableMetadata As DbBaseTableMetadataModel = New DbBaseTableMetadataModel("Newsletter Ad Banner Layouts", "NewsletterAdBannerLayouts", "default", False)
        '
        Public Property rowcnt As Integer
        Public Property columncnt As Integer
        Public Property pxcolumnspace As Integer
        Public Property pxrowspace As Integer
        '
    End Class
End Namespace

