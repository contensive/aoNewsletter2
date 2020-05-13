Option Explicit On
Option Strict On

Imports Contensive.BaseClasses


Namespace Controllers
    Public Class DesignBlockController
        '
        '====================================================================================================
        ''' <summary>
        ''' return the instanceId for a design block. It should be an document argument set when the addon is dropped on the page.
        ''' If the addon is created with a json string it should be included as an argument
        ''' If it is not included, the page id is used to make a string
        ''' If no instanceId can be created a blank is returned and should NOT be used.
        ''' If returnHtmlMessage is non-blank, add it to the html
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <param name="designBlockName">A name of the design block. This must be unqiue for each type of design block (i.e. text, tile, etc)</param>
        ''' <param name="returnHtmlMessage"></param>
        ''' <returns>If blank is returned, </returns>
        Public Shared Function getSettingsGuid(cp As CPBaseClass, designBlockName As String, ByRef returnHtmlMessage As String) As String
            '
            ' -- check arguments
            If (String.IsNullOrWhiteSpace(designBlockName)) Then Throw New ApplicationException("getInstanceId called without valid designBlockName.")
            '
            ' -- if this code is running during page rendering, read the instanceId from the process
            Dim result As String = cp.Doc.GetText("instanceId")
            If (Not String.IsNullOrWhiteSpace(result)) Then Return result
            '
            ' -- if this code is running as a remoteMethod, read the instanceId from the form submission
            result = cp.Doc.GetText("forminstanceId")
            If (Not String.IsNullOrWhiteSpace(result)) Then Return result
            '
            ' -- if there is no instanceId added to the rendering context, try the page Id
            If (cp.Doc.PageId > 0) Then
                '
                ' -- no instance Id, create a unquie string for this page, but display error is already used on this page
                result = "DesignBlockUsedWithoutInstanceId-[" & designBlockName & "]-PageId-" & cp.Doc.PageId.ToString()
                If (Not String.IsNullOrEmpty(cp.Doc.GetText(result))) Then
                    '
                    ' -- no instance Id, second occurance, display error
                    returnHtmlMessage &= "<p>Error, this design block is used twice on this page. This is only allowed if it was added with the drag-drop tool, or includes a unique instance id.</p>"
                    cp.Site.ErrorReport("Design Block [" & designBlockName & "] on page [#" & cp.Doc.PageId & "," & cp.Doc.PageName & "] does not include an instanceId and was used on the page twice. This is not allowed. To use it twice, used the drag-drop design block tool or manually add the argument ""instanceid"" : ""{unique-guid}"".")
                    Return String.Empty
                End If
                cp.Doc.SetProperty(result, "used")
                Return result
            End If
            '
            ' -- if this is the admin site, make one up
            If (cp.Request.PathPage = cp.Site.GetText("adminurl")) Then
                '
                ' -- addon run on admin site
                result = "DesignBlockUsedOnAdminSite-[" & designBlockName & "]"
                If (Not String.IsNullOrEmpty(cp.Doc.GetText(result))) Then
                    '
                    ' -- admin site, second occurance, display error
                    returnHtmlMessage &= "<p>Error, this design block is used twice on the admin site. This is only allowed if it was added with the drag-drop tool, or includes a unique instance id.</p>"
                    cp.Site.ErrorReport("Design Block [" & designBlockName & "] on page [#" & cp.Doc.PageId & "," & cp.Doc.PageName & "] does not include an instanceId and was used on the page twice. This is not allowed. To use it twice, used the drag-drop design block tool or manually add the argument ""instanceid"" : ""{unique-guid}"".")
                    Return String.Empty
                End If
                Return result
            End If
            Throw New ApplicationException("Design Block [" & designBlockName & "] called without instanceId must be on a page or the admin site.")
        End Function
    End Class
End Namespace

