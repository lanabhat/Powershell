<#
.Synopsis
   Script to convert a input object to html format
.DESCRIPTION
   This script converts an object which has a flat formatting to a searchable and sortable HTML datatable using an HTML template. If email ids are supplied script will email the report
.EXAMPLE
  .\Create-HTMLDataTable.ps1 -InputObject (@{Title="testReport";Content=(get-service | select Name,Status,DisplayName)}) -ReportTitle "Test report" -ShowReport -templatePath .\Templates\DataTableTemplate_BLUE_Online.html  -ReportSavePath .\MyReports.EXAMPLE
   .\Create-HTMLDataTable.ps1 -InputObject (Import-Csv .\csvName.csv) -ReportTitle "Test CSV Import" -ReportSavePath $env:userprofile\desktop -ShowReport -emailIDs "lakshminarayana.govinda@socgen.com","BLR-BLR-GTS-EUS-DVS-RDS-DEL@socgen.com"
.INPUTS
   None
.PARAMETER InputObject
    An object of any type which is not a nested object.Example a result of import csv.
.PARAMETER ReportTitle
    Title to be used for the report, it will be added to the HTML and also used as the file name
.PARAMETER ReportSavePath
    Path to save the report file. Defaults to users desktop.
.PARAMETER ShowReport
    Launch the report file in default browser after completing script execution
.PARAMETER TemplatePath
    HTML Template file path, template file has the formatting required to generate the required results
    default path taken is .\Templates\
.PARAMETER EmailIds
    Email ids to send the report after completion
.OUTPUTS
   Generates HTML report file   
.NOTES
   
.COMPONENT
   
.ROLE
   
.FUNCTIONALITY
   HTML datatable report generation
#>

[CmdletBinding(SupportsShouldProcess = $true, 
    PositionalBinding = $false,                  
    ConfirmImpact = 'Medium')]    
Param
(
    #Input object which will be sent as a table
    [Parameter(Mandatory = $true, 
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true, 
        ValueFromRemainingArguments = $false,                    
        Position = 0)]
    [ValidateScript( {$_.Title -ne $null -and $_.Title.length -gt 4})]
    [Alias("TableObj")] 
    [psobject[]]
    $InputObject,

    # Title of the report file
    [Parameter(Mandatory = $true, 
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true, 
        ValueFromRemainingArguments = $false, 
        Position = 1)]       
    [ValidateNotNullOrEmpty()]
    [string]
    $ReportTitle = "HTML_Report",

    # Path to save the report file to
    [Parameter(Mandatory = $false, 
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true, 
        ValueFromRemainingArguments = $false, 
        Position = 2)]    
    [ValidateScript( {Test-Path $_})]          
    [String]
    $ReportSavePath = "$($env:USERPROFILE)\desktop",

    #Show the report on screen after generation
    [Parameter(Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true,
        Position = 3
    )]
    [SWITCH]$ShowReport,

    #HTML Template path
    [Parameter(Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true,
        Position = 4
    )]
    $templatePath = "$PSScriptRoot\Templates\DataTableTemplate_Online.html",
    #Email IDs to emiail the report to
    [Parameter(Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true,
        Position = 4       
    )]
    [string[]]
    $emailIDs      
)

Begin {
    Function Send-Email {
        param($sub, $msg, $attachment, $emailIds)
       
        $Bodyout = $msg
        $messageParameters = @{    
                     
            Subject    = $sub
            Body       = ConvertTo-HTML -Body "$Bodyout" -Title "$sub $(get-date -format dd/MM/yyyy)" | out-string
            from       = 'myEmail@email.com'  
            To         = $emailIds              
            SmtpServer = "smtpserver.Address.com"
              
        }
        if ($attachment) {
            Send-MailMessage @messageParameters -BodyAsHtml -Attachments $attachment
        }
        else {
            Send-MailMessage @messageParameters -BodyAsHtml #-Attachments $report
        }
    }

    Write-Verbose "Setting variables"
    $scriptpath = $PSScriptRoot
    $reportName = "$($ReportTitle.replace(" ","_"))_$(Get-Date -UFormat "%Y_%m_%d").htm"
    $finalHTML = ""
          
    $reusableTemplate = @"     
     <script>
    `$(document).ready(function() {
    `$('#ElementID').DataTable( {
        "lengthMenu": [[10, 25, 50, -1], [10, 25, 50, "All"]]
    } );
    } );
    </script>           
        <div class="mysection">
        <H3><div class="myheader"> #TABLE TITLE#</div></H3>              
        <table id="ElementID" class="display" width="100%" cellspacing="0">#PUT YOUR TABLE HERE#</table>
        </div>
"@

    #$reportName = $ReportTitle

    Write-Verbose "Report save path: $ReportSavePath"        
    Write-Verbose "Report title: $ReportTitle"
    Write-Verbose "Report name: $reportName"
    Write-Verbose "Show report: $ShowReport"
    
    $IDCount = 0
    if (!(Test-Path $templatePath)) {      
        $defaultTemplate = "$scriptpath\Templates\DataTableTemplate_Online.html"
        Write-Verbose "Default template path: $defaultTemplate"
        if (Test-Path $defaultTemplate) {$templatePath = $defaultTemplate; Write-Warning "Template not found ! Set to default template $templatePath"}
        else {Write-Error "Template not found"; return "Template file missing $defaultTemplate"}      
    }
    else {
        try {
            Write-Verbose "Validating report template file"
            if (!((Get-Content $templatePath) -match "#BODY#")) {
                Write-Error "Template format is not correct, please check the template $templatePath"
                return $false

            }
        }
        catch {
            Write-Error "Error while reading template $templatePath"        
            $_
        }
    }  
    $reusableTemplate = $reusableTemplate.replace("#REPORT CREATED#", (Get-Date))
}
Process {       
       
    Write-Verbose "Processing inputs"
    foreach ($obj in $InputObject) {
        $IDCount ++
        $tableTitle = $obj.Title
        $reportContent = $obj.Content
         
        #Issu passing single object, it is being considered as an array, added title to it         
        
        $rawHtmlString = $reportContent | ConvertTo-Html -Fragment
        $htmlString = $rawHtmlString | % {$_.replace("<table>", "")} | % {$_.replace("</table>", "")} | % {$_.Replace("<tr><th>", "<thead><tr><th>")} | % {$_.Replace("</th></tr>", "</th></tr></thead>")}        
         
        if (-not $reportContent) {$htmlString = "<TR><TH>Details<TH></TR></TR><TD><center>No objects available</center></TD></TR>"}
        Write-Verbose "Generating report from template"
        
        $newHtml = $reusableTemplate.replace("#TABLE TITLE#", $tableTitle)
        $newHtml = $newHtml.replace("ElementID", "ItemID_$IDCount")         
        $newHtml = $newHtml.replace("#PUT YOUR TABLE HERE#", $htmlString)
        $newHtml = "$newHtml<P></P>"
        Write-Verbose "Writing report file to disk"

        $finalHTML += $newHtml          
    }
       
}
End {
    $htmlFileCont = Get-Content -ReadCount 0 -Path $templatePath 
    #Add the below line to HTML Template file   
    #<Div class="ReportCreated">Report created on: #REPORT CREATED#</Div>  
    $htmlFileCont = $htmlFileCont.replace("#REPORT TITLE#", "$ReportTitle $(Get-Date -UFormat "%d-%m-%Y")")
    $htmlFileCont = $htmlFileCont.replace("#REPORT CREATED#", $(Get-Date -UFormat "%d-%m-%Y"))
    $htmlFileCont = $htmlFileCont.replace("#BODY#", $finalHTML)

    $htmlFileCont | Out-File $ReportSavePath\$reportName -Verbose -Force
    Write-Host "Report saved to: $ReportSavePath" -ForegroundColor Green
    if ($ShowReport) {
        Write-Verbose "Launching report"
        Invoke-Item $ReportSavePath\$reportName
    }
    if ($emailIDs) {
        Write-Verbose "Emailing report"
        Send-Email -sub "HTML Report for $ReportTitle Date: $(Get-Date -UFormat "%d-%m-%Y")" -msg "Please find the attachment" -attachment $ReportSavePath\$reportName -EmailIds $emailIDs      
        Write-Verbose "Email sent"
    }   

}


