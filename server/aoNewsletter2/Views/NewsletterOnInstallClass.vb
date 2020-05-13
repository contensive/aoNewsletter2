
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
    Public Class NewsletterOnInstallClass
        Inherits AddonBaseClass
        '
        '====================================================================================================
        '
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Try
                '
                ' -- read instanceId, guid created uniquely for this instance of the addon on a page
                Dim versionLastInstalled As Integer = CP.Site.GetInteger(spName_NewsletterVersion, 0)
                If (newsletterVersion > versionLastInstalled) Then
                    '
                    ' -- upgraded needed
                    If (versionLastInstalled <= 3) Then
                        '
                        ' -- remove old settings addon (needed until deprecated property is included in metadata)
                        CP.Db.ExecuteNonQuery("delete from ccaggregatefunctions where ccguid='{fa787411-f505-433d-990b-47bb55473ef0}'")
                    End If

                End If
                    Return String.Empty
            Catch ex As Exception
                CP.Site.ErrorReport(ex)
                Return "<!-- Unexpected Exception -->"
            End Try
        End Function
    End Class
End Namespace