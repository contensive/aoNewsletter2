﻿
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace newsletter2
    '
    ' Sample Vb addon
    '
    Public Class newsletterExtensionClass
        Inherits AddonBaseClass
        '
        ' - update references to your installed version of cpBase
        ' - Verify project root name space is empty
        ' - Change the namespace to the collection name
        ' - Change this class name to the addon name
        ' - Create a Contensive Addon record with the namespace apCollectionName.ad
        ' - add reference to CPBase.DLL, typically installed in c:\program files\kma\contensive\
        '
        '=====================================================================================
        ' 
        '=====================================================================================
        '
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim returnHtml As String = ""
            Try
                returnHtml = "Visual Studio Contensive Addon - OK response"
            Catch ex As Exception
                errorReport(CP, ex, "execute")
            End Try
            Return returnHtml
        End Function
        '
        '=====================================================================================
        ' common report for this class
        '=====================================================================================
        '
        Private Sub handleError(ByVal cp As CPBaseClass, ByVal ex As Exception, ByVal method As String)
            Try
                cp.Site.ErrorReport(ex, "Unexpected error in newsletterExtensionClass." & method)
            Catch exLost As Exception
                '
                ' stop anything thrown from cp errorReport
                '
            End Try
        End Sub
        '
        '===========================================================================================================
        '   Tag Type
        '       Issue - The copy is the same for each issue of the newsletter
        '       Page - The copy is new for each page
        '===========================================================================================================
        '
        Public Function GetContent(cp As CPBaseClass, OptionString As String) As String
            'On Error GoTo ErrorTrap
            '

            Dim Stream As String
            Dim ExtensionName As String
            Dim ExtensionType As String
            Dim Copy As String
            Dim PageID As Integer
            Dim IssueID As Integer
            Dim IsWorkflowRendering As Boolean
            Dim IsQuickEditing As Boolean
            Dim Common As New CommonClass
            Dim NewsletterProperty As String
            Dim Parts() As String
            Dim NewsletterID As Integer
            '
            If Not (Main Is Nothing) Then
                '
                ' Assume newsletterNavClass is used within a PageClass
                ' Get the Issue and Newsletter from the visit properties set in PageClass
                '
                ExtensionName = Trim(cp.Doc.GetText("ExtensionName", OptionString))
                NewsletterProperty = Main.GetVisitProperty(VisitPropertyNewsletter)
                Parts() = Split(NewsletterProperty, ".")
                If UBound(Parts) > 2 Then
                    NewsletterID = kmaEncodeInteger(Parts(0))
                    IssueID = kmaEncodeInteger(Parts(1))
                    PageID = kmaEncodeInteger(Parts(2))
                    'FormID = kmaEncodeInteger(Parts(3))
                End If
                If ExtensionName = "" Then
                    ExtensionName = "Default"
                    'If Main.IsAdmin() Then
                    '    GetContent = common.getAdminHintWrapper( cp,"The ExtensionName is blank. To use the Page Extension, set the ExtensionName and select the ExtensionType.")
                    'End If
                Else
                    '
                    ' Handle PageID Request Variable
                    '
                    ExtensionType = LCase(Trim(cp.Doc.GetText("ExtensionType", OptionString)))
                    Call cp.Site.TestPoint("GetIssueID call 1, NewsletterID=" & NewsletterID)
                    IssueID = Common.GetIssueID(cp, NewsletterID)
                    PageID = cp.doc.getInteger(RequestNameIssuePageID)
                    IsQuickEditing = Main.IsQuickEditing("Page Content")
                    IsWorkflowRendering = Main.IsWorkflowRendering
                    '
                    Select Case ExtensionType
                        Case "issue"
                            If IssueID <> 0 Then
                                GetContent = Main.GetContentCopy("Newsletter-Extension-Issue-" & IssueID & "-" & ExtensionName)
                            End If
                        Case "page"
                            If PageID <> 0 Then
                                GetContent = Main.GetContentCopy("Newsletter-Extension-Issue-Page-" & IssueID & "-" & PageID & "-" & ExtensionName)
                            End If
                        Case Else
                            GetContent = Common.getAdminHintWrapper(cp, "The Extension Type is blank. To use the Page Extension, set the ExtensionName and select the ExtensionType.")
                    End Select
                End If
            End If
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError("ExtensionClass", "GetContent")
        End Function

    End Class
End Namespace
