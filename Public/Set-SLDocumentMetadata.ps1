Function Set-SLDocumentMetadata  {


    <#

.SYNOPSIS
    Set document metadata.

.DESCRIPTION
    Set document metadata that helps identify a document and also to organise them by tags,comment or author.


.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER Title
    The title of the document.

.PARAMETER Author
    The creator of the document.

.PARAMETER Comment
    The summary or abstract of the contents of the document.

.PARAMETER Tags
    A word or set of words describing the document.Refers to keywords in excel.

.PARAMETER Category
    The category of the document eg: personal,business,financial etc.

.PARAMETER LastModifiedBy
    The document is last modified by this person.

.PARAMETER Subject
    The topic of the document.


.Example
    PS C:\> Get-SLDocument C:\temp\test.xlsx | Set-SLDocumentMetadata -Title mydoc -Author kiran -Comment "this is a test doc" -Tags "test;document" | Save-SLDocument


    Description
    -----------
    Set the title,author,comment and tags properties on the document named test.
    Note:Tags are seperated by semicolons.


.INPUTS
   String,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    N/A

#>


    #>
    [CmdletBinding()]
    [OutputType([SpreadsheetLight.SLDocument])]
    param (
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [parameter(Mandatory = $false)]
        [string]$Title,

        [parameter(Mandatory = $false)]
        [string]$Author,

        [parameter(Mandatory = $false)]
        [string]$Comment,

        [parameter(Mandatory = $false)]
        [string]$Tags,

        [parameter(Mandatory = $false)]
        [string]$Category,

        [parameter(Mandatory = $false)]
        [string]$LastModifiedBy,

        [parameter(Mandatory = $false)]
        [string]$Subject
    )
    PROCESS
    {
        if ($Author) { $WorkBookInstance.DocumentProperties.Creator = $Author }
        if ($Title) { $WorkBookInstance.DocumentProperties.Title = $Title }
        if ($Comment) { $WorkBookInstance.DocumentProperties.Description = $Comment }
        if ($Tags) { $WorkBookInstance.DocumentProperties.Keywords = $Tags }
        if ($Category) { $WorkBookInstance.DocumentProperties.Category = $Category }
        if ($LastModifiedBy) { $WorkBookInstance.DocumentProperties.LastModifiedBy = $LastModifiedBy }
        if ($Subject) { $WorkBookInstance.DocumentProperties.Subject = $Subject }

        Write-Output $WorkBookInstance
    }

}
