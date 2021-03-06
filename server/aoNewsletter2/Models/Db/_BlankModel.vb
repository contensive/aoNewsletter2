﻿Imports System
Imports Contensive.Addons.Newsletter.Models.Db
Imports Contensive.BaseClasses
Imports Contensive.Models.Db

Namespace Models.Db
    Public Class BlankModel
        Inherits Global.Contensive.Models.Db.DbBaseModel

        Public Shared ReadOnly Property tableMetadata As DbBaseTableMetadataModel = New DbBaseTableMetadataModel("Newsletter Ad Banner Layouts", "NewsletterAdBannerLayouts", "default", False)
        '
        Public Property imageFilename As String
        '
    End Class
End Namespace

