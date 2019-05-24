# ***** Variable Set *****

param(
[Parameter(Mandatory=$true)][String]$runSPIChoice='',
[Parameter(Mandatory=$true)][boolean]$runHTTPSPostTest, 
[Parameter(Mandatory=$true)][boolean]$runHTTPPostTest, 
[Parameter(Mandatory=$true)][boolean]$runPrintTest='',
[Parameter(Mandatory=$true)][boolean]$runNetworkShareTest='',
[Parameter(Mandatory=$true)][boolean]$runPrecisionCheck_Numeric='',
[Parameter(Mandatory=$true)][String]$networkFolder=''
)

Write-Host "*** NetworkFolder Recieved: $networkFolder"
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[int]$Count = 200      #  number to generate
[int]$P = 1           #  progress count holder
[int]$b = 1           #  test number incrimenter
[int]$p = 1           #  data array pointer
[int]$g = 1           #  group array pointer
[int]$tot = 1         #  total test counter
[string]$ePOServer = "Korea"
[string]$to = "Destination Address"
[string]$frn = "sender Address"
[string]$CC = "CC Address"
[string]$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
  #Write-Host "scriptPath = $scriptPath"
[string]$WatinPath = "$scriptPath\lib\Watin\bin\net40\Watin.Core.dll" 
[string]$TestFileFolder = "$scriptPath\TestFiles\"
[string]$OutputFolder = "$scriptPath\OutputFolder\"
[string]$LogFolder = "$scriptPath\Log\"
  #Write-Host "OutputFolder = $OutputFolder"
[string]$LogName = "Log.csv"
[string]$LogPath = "$LogFolder$LogName"
[string]$Host = HostName
[string]$scriptStartTime = Get-Date
[string]$scriptStartTimeUTC = Get-Date -Format FileDateTimeUniversal
[string]$FTSPre= "DLP-TEST"
[string]$AFTPrefix = "DLP-AFT-TEST_"
[string]$AFTVersion = "2-3"
[string]$AFTName = "$AFTPrefix$AFTVersion"  + "_{" + $Host + "_" + $scriptStartTimeUTC + "}" 
[string]$filePath = ""
$NumOfFiles = 0
[string]$SPISelection_NewName = ""
[string]$SPISelection_NewPath = ""

# update log when action is taken
function UpDate-Log ($CSV_ActionType, $CSV_SPI){

 $CurrentTime=Get-Date
 $LogEntry = "$scriptStartTime, $Host, $CSV_ActionType, $CSV_SPI, $CurrentTime"
 Add-Content -Path $LogPath  -Value $LogEntry

}

# temporarily change file name for unique testing ID

if ($runSPIChoice -eq "-- ALL SPI --") {

# change all file names and set counter to dynamically count

$count = 1

Get-Childitem -Path "$PSScriptRoot\TestFiles" -File -Filter "*.txt" |
    ForEach-Object { 
        $NumOfFiles = $NumOfFiles + 1
        $filePath = $_.FullName
        $newName = $_.Name -replace $FTSPre, $AFTName

        Rename-Item -LiteralPath $filePath $newName  
        }

}
Else {
# change only one file name and set counter to static 1
    $filePath = "$PSScriptRoot\TestFiles\$runSPIChoice"
    $newName = $runSPIChoice -replace $FTSPre, $AFTName
    Rename-Item -LiteralPath $filePath $newName 
    $SPISelection_NewName = $newName
    $SPISelection_NewPath = "$PSScriptRoot\TestFiles\$newName" 

    $count = 1
    $NumOfFiles = 1
}

#array sets and counters
#$SPI = @("Peoples Republic of China ID","UK NIN ID","UK Sort Code","CC - Visa Number")

#dynamic list of test files to execute
#if ($runSPIChoice -eq "-- ALL SPI --"){
##all the stuff
#}

Write-Host "
*************************
Beginning Automated Test
*************************

IMPORTANT MESSAGE BELOW:

PHASE I: Hands-Off [10 - 20 minutes]

* Abstain from using your computer for approximately 10 minutes while the Auatomated Web Posting is executed.

PHASE II: Run in Background [1 hour]

* Once the Print testing begins, you may resume use of your computer. The print test may take up to 1 hour.

This is a ""hands off"" proccess. Please let this, and windows that appear, run in the back ground uninterrupted.

This window will automatically close once all testing is done.
The expected total run time is approximately 1 hour and 10 minutes.
"

 # ******  Web Post     *****
# check for any web posting selection. initialize common components, and proceed.

if ($runHTTPSPostTest -eq $True -Or $runHTTPPostTest -eq $True) {

    $WatinPath = "$scriptPath\lib\Watin\bin\net40\Watin.Core.dll" 
    $TestFileFolder = "$scriptPath\TestFiles\"
    $watin = [Reflection.Assembly]::LoadFrom($WatinPath)

    function WebPost-File
    {
    param(
        [Parameter(Mandatory=$true)][String]$FileLocation='',
        [Parameter(Mandatory=$true)][String]$webAddress='',
        [Parameter(Mandatory=$true)][String]$browseButton='',
        [Parameter(Mandatory=$true)][String]$submitButton=''
            )
    $ie = new-object WatiN.Core.IE($webAddress)
    $ie.Visible = $true
    $ie.WaitForComplete()
    $file1 = $ie.FileUpload($browseButton) #id of the input
    write-host "|| Posting the Following File at $webAddress ||
$TestFileFolder$FileLocation
"
    $file1.set("$TestFileFolder$FileLocation")
    $o = $ie.Button($submitButton) #id of the submit button
    $o.Click()
    $ie.WaitForComplete()
    #$ie.GoTo("https://dlptest.com/https-post/")
    #$ie.WaitForComplete()
    #$ie.Refresh()
    $ie.Close()
    $ie.Dispose()
    UpDate-Log $webAddress $FileLocation
    }
    

    # ****** HTTPS Web Post     *****

    #check for run https post test param arg
    If ($runHTTPSPostTest -eq $True) {
    
        # set HTTPS specific web page interacion variables
        $webAddress="https://dlptest.com/https-post/"
        $browseButton="wpforms-513-field_3"
        $submitButton="wpforms-submit-513"

        Write-Host "
        _______________________________
               PHASE I: Hands-Off

         -- Running HTTPS Post Test --

        Estimated Time: 10 - 20 minutes
        _______________________________


        * This is a ""hands off"" proccess. Please let this, and windows that appear, run in the back ground uninterrupted.

        * During this automated Post test, it is advised that you abstain from interacting with your computer, as this may interfere with the automted submissions.

        * This process should take less than 10 minutes.



        "


        if ($runSPIChoice -eq "-- ALL SPI --") {

        $count = 1
        #$NumOfFiles = 0

        # Counter for files
        #Get-Childitem -Path "$PSScriptRoot\TestFiles" -File -Filter "*.txt" |
        #    ForEach-Object { $NumOfFiles = $NumOfFiles + 1 }

        # Calls Web Post function for each file
        Get-Childitem -Path "$PSScriptRoot\TestFiles" -File -Filter "*.txt" |
            ForEach-Object { 
        
                Write-Host "File $count / $NumOfFiles" 
    
                WebPost-File $_ $webAddress $browseButton $submitButton
                #UpDate-Log "HTTPS_Post" $_
        
                $count = $count + 1
        
                }

        $count = $Count - 1
        }
        Else {
        #PostFile $runSPIChoice
        WebPost-File $SPISelection_NewName $webAddress $browseButton $submitButton
        $count = 1
        #$NumOfFiles = 1
        }
        Write-Host "
        _____________________________
              PHASE I: Hands-Off

        -- Closing HTTPS POST Test --

        $count / $NumOfFiles Files Handled
        _____________________________
        "
        }

    # ****** HTTP Web Post     *****

    #check for run post test param arg
    If ($runHTTPPostTest -eq $True) {

        $WatinPath = "$scriptPath\lib\Watin\bin\net40\Watin.Core.dll" 
        $TestFileFolder = "$scriptPath\TestFiles\"
          #Write-Host "TestFileFolder = $TestFileFolder"
        $OutputFolder = "$scriptPath\OutputFolder\"
          #Write-Host "OutputFolder = $OutputFolder"

        #$WatinPath = 'C:\Users\marksimm\WatiN\bin\net40\WatiN.Core.dll' #path with downloaded assembly
        $watin     = [Reflection.Assembly]::LoadFrom( $WatinPath )

        #$ie        = new-object WatiN.Core.IE("https://dlptest.com/https-post/")
        #$ie.WaitForComplete()
        #$ie.Visible = $false
        #$file1 = $ie.FileUpload('wpforms-513-field_3') #id of the input

        # open and close of object must be in loop to avoid execution errors
        function PostFile
        {
            param($FileLocation)
            $ie        = new-object WatiN.Core.IE("https://dlptest.com/http-post/")
            #$ie        = New-Object WatiN.Core.IE("http://10.95.113.13/ngdc2/dlpfileuploadtest.html")
            #$ie        = new-object WatiN.Core.IE
            $ie.Visible = $false
            #$ie.GoTo("https://dlptest.com/https-post/")
            $ie.WaitForComplete()
            $file1 = $ie.FileUpload('wpforms-518-field_3') #id of the input
            #$file1 = $ie.FileUpload('FILE1') #id of the input
            write-host "|| Posting the Following File at https://dlptest.com/http-post/ ||"
            #write-host "|| Posting the Following File at http://10.95.113.13/ngdc2/dlpfileuploadtest.html ||"
            write-host "$TestFileFolder$FileLocation"
            write-host ""
            $file1.set("$TestFileFolder$FileLocation")
            $o = $ie.Button('wpforms-submit-518') #id of the submit button
            #$o = $ie.Button('Upload!') #id of the submit button
            $o.Click()
            $ie.WaitForComplete()
            #$ie.GoTo("https://dlptest.com/https-post/")
            #$ie.WaitForComplete()
            #$ie.Refresh()
            $ie.Close()
            $ie.Dispose()
            UpDate-Log "Post_HTTP" $FileLocation
        }

    Write-Host "
    ______________________________
            PHASE I: Hands-Off

    -- Running HTTP POST Test --

        Estimated Time: 10-20 Minutes
    ______________________________


    * This is a ""hands off"" proccess. Please let this, and windows that appear, run in the back ground uninterrupted.

    * During this automated Post test, it is advised that you abstain from interacting with your computer, as this may interfere with the automated submissions.

    * This process should take less than 10 minutes.



    "


    if ($runSPIChoice -eq "-- ALL SPI --") {

    $count = 1
    #$NumOfFiles = 0

    # Counter for files
    #Get-Childitem -Path "$PSScriptRoot\TestFiles" -File -Filter "*.txt" |
    #    ForEach-Object { $NumOfFiles = $NumOfFiles + 1 }

    # Calls man Web Post function
    Get-Childitem -Path "$PSScriptRoot\TestFiles" -File -Filter "*.txt" |
        ForEach-Object { 
        
            Write-Host "File $count / $NumOfFiles" 
    
            PostFile $_
            #UpDate-Log "HTTPS_Post" $_
        
            $count = $count + 1
        
            }

    $count = $Count - 1
    }
    Else {
    #PostFile $runSPIChoice
    PostFile $SPISelection_NewName
    $count = 1
    #$NumOfFiles = 1
    }
    Write-Host "
    _____________________________
            PHASE I: Hands-Off

    -- Closing HTTP Post Test --

    $count / $NumOfFiles Files Handled
    _____________________________
    "
    }

}

# ****** Print to PDF     *****

#check for print to PDF param arg
if ($runPrintTest -eq $True) {

$PrintPageHandler =
{
    param([object]$sender, [System.Drawing.Printing.PrintPageEventArgs]$ev)

    $linesPerPage = 0
    $yPos = 0
    $count = 0
    $leftMargin = $ev.MarginBounds.Left
    $topMargin = $ev.MarginBounds.Top
    $line = $null

    $printFont = New-Object System.Drawing.Font "Arial", 10

    # Calculate the number of lines per page.
    $linesPerPage = $ev.MarginBounds.Height / $printFont.GetHeight($ev.Graphics)

    # Print each line of the file.
    while ($count -lt $linesPerPage -and (($line = $streamToPrint.ReadLine()) -ne $null))
    {
        $yPos = $topMargin + ($count * $printFont.GetHeight($ev.Graphics))
        $ev.Graphics.DrawString($line, $printFont, [System.Drawing.Brushes]::Black, $leftMargin, $yPos, (New-Object System.Drawing.StringFormat))
        $count++
    }

    # If more lines exist, print another page.
    if ($line -ne $null) 
    {
        $ev.HasMorePages = $true
    }
    else
    {
        $ev.HasMorePages = $false
    }
}

function Out-Pdf
{
    param($InputDocument)

    write-host $InputDocument
    Add-Type -AssemblyName System.Drawing
    Write-Host "|| Printing the Following File to PDF ||"
    Write-Host "$PSScriptRoot\TestFiles\$InputDocument"
    Write-Host ""
    $doc = New-Object System.Drawing.Printing.PrintDocument
    $doc.DocumentName = $InputDocument.FullName
    $doc.PrinterSettings = New-Object System.Drawing.Printing.PrinterSettings
    $doc.PrinterSettings.PrinterName = 'Microsoft Print to PDF'
    $doc.PrinterSettings.PrintToFile = $true

    $streamToPrint = New-Object System.IO.StreamReader $InputDocument.FullName

    $doc.add_PrintPage($PrintPageHandler)

    $doc.PrinterSettings.PrintFileName = "$($InputDocument.DirectoryName)\$($InputDocument.BaseName).pdf"
    $doc.Print()

    $streamToPrint.Close()
    UpDate-Log "Print_ToPDF" $InputDocument
}


Write-Host "
______________________________
 PHASE II: Run in Background

  -- Running Print Test --

    Estimmated Time: 1 hr
______________________________


* This is a ""hands off"" proccess. Please let this, and windows that appear, run in the back ground uninterrupted.

* During this automated Print test, you may resume use of your machine. This may run in the background for approximately 1 hour.



"

if ($runSPIChoice -eq "-- ALL SPI --") {

	$count = 1
	# $NumOfFiles = 0

	# Get-Childitem -Path "$PSScriptRoot\TestFiles" -File -Filter "*.txt" |
	#	 ForEach-Object { $NumOfFiles = $NumOfFiles + 1 }

	Get-Childitem -Path "$PSScriptRoot\TestFiles" -File -Filter "*.txt" |
		ForEach-Object { 
			
			Write-Host "File $count / $NumOfFiles" 
			Out-Pdf $_ 
			$count = $count + 1
			 
			 }

	$count = $Count - 1
}
Else {
			
			Out-Pdf $SPISelection_NewPath 
			 
	$count = 1
	# $numOfFiles = 1
} 

	#Write-Host ""
	Write-Host "
______________________________
 PHASE II: Run in Background

   -- Closing Print Test --

$count / $NumOfFiles Files Handled
______________________________

"
}

#$SPISelection_NewPath
# ****** Network Share     *****

function NetworkShare {
        #param($FileLocation) #put in additional param for source path

        param(
            [Parameter(Mandatory=$true)][String]$FileName='',
            [Parameter(Mandatory=$true)][String]$SourcePath=''
            )

        #Write-host "

        #$FileLocation
        #$driveRoot

        #"

        Write-Host "|| Copying the Following File to Your Network Drive ||"
        Write-Host "$FileName
        "
        #Write-Host "Destination: $NetworkShareFolder"

        Copy-Item -LiteralPath "$SourcePath\$FileName" $NetworkShareTestFolder
        #Copy-Item $TestFileFolder$FileLocation $NetworkShareTestFolder
        UpDate-Log "Network_Share" $FileLocation

        $filePath = "$NetworkShareTestFolder\$FileName"
        #$oldName = $FileLocation -replace $AFTName, $FTSPre
        Write-Host "Debug: Remove-Item file path: 
        $filePath"
        Remove-Item -LiteralPath $filePath      
    }

#write-host "

#runPrecisionCheck_Numeric
#$runPrecisionCheck_Numeric

#"
#check for run post test param arg
If ($runNetworkShareTest -eq $True) {

    $TestFileFolder = "$scriptPath\TestFiles\"

    #Get homedrive location, and create test file destination.

    $NetworkShareTestFolder = "$networkFolder\DLP-AFT_TEST"
    New-Item -Path $NetworkShareTestFolder  -ItemType directory -Force
    
    #display console update
    Write-Host "
    _________________________________
       PHASE II: Run in Background

    -- Running Network Share Test --

       Estimated Time: 10 Minutes
    _________________________________


    * During this portion of the test, you may interact with your computer.

    * This will attempt to copy the selected file(s) to your network home drive.

    * This process should take less than 10 minutes.



    "

    ## Check SPI Choice and run appropraite routine
    if ($runSPIChoice -eq "-- ALL SPI --") {
    #run "all"' choice process
    $count = 1
    #$NumOfFiles = 0

    Get-Childitem -Path "$PSScriptRoot\TestFiles" -File -Filter "*.txt" |
        ForEach-Object { 
        
            Write-Host "File $count / $NumOfFiles" 
            #$Name = $_.Name
    
            NetworkShare $_.Name $_.DirectoryName
            #UpDate-Log "HTTPS_Post" $_


        
            $count = $count + 1
        
            }

    $count = $Count - 1
    }

    Else {
    #run single file process
    ## ** look for new name
    Write-Host "File $count / $NumOfFiles" 
    NetworkShare $SPISelection_NewName "$PSScriptRoot\TestFiles"
    $count = 1
    $NumOfFiles = 1
    }

    #delete the contents of the folder

    Write-Host "
    _________________________________
      PHASE II: Run in Background

    -- Closing Network Share Test --

    $count / $NumOfFiles Files Handled
    _________________________________
    "
}

## ********* Numeric Precision Search ************

If ($runPrecisionCheck_Numeric -eq $True) {

Write-Host "
______________________________________
     PHASE II: Run in Background

-- Running Numeric Precision Search --
    
      Estimated Time: 10 Minutes
______________________________________


    * During this portion of the test, you may interact with your computer.

    * This will attempt to copy the selected file(s) to your network home drive.

    * This process should take less than 10 minutes.


"
    if ($runSPIChoice -eq "-- ALL SPI --") {

        # change all file names and set counter to dynamically count

        $count = 1

        $NumOfFiles = 0
        Get-Childitem -Path "$PSScriptRoot\TestFiles\NumericPrecisionSearch" -File -Filter "*.txt" |
            ForEach-Object { 
                $NumOfFiles = $NumOfFiles + 1
                $filePath = $_.FullName
                $newName = $_.Name -replace $FTSPre, $AFTName

                Rename-Item -LiteralPath $filePath $newName  

                }

            $count = 1
            # $NumOfFiles = 0

        # iniate transfer files to network share
        Get-Childitem -Path "$PSScriptRoot\TestFiles\NumericPrecisionSearch" -File -Filter "*.txt" |
            ForEach-Object { 
             
                Write-Host "File $count / $NumOfFiles" 

                NetworkShare $_.Name $_.DirectoryName
                UpDate-Log "Precision_Numeric" $_

                $count = $count + 1
                ##*** rename file back to original
        
                }

            $count = $Count - 1

    }

    Else {
    ## GET DATACLASS NAME FROM SPI SELECTED ON RUN
    $dataClass = "_" + $runSPIChoice.Substring($runSPIChoice.IndexOf("TEST_") + 4, $runSPIChoice.IndexOf("[") - $runSPIChoice.IndexOf("TEST_") -5)
    ## MATCH ELEMENT NAME TO PRECISION SEARCH FILE AND EXECUTE SHARE
Get-Childitem -Path "$PSScriptRoot\TestFiles\NumericPrecisionSearch" -File -Filter "*.txt" |
            ForEach-Object {
            $Name = $_.Name
            if ($Name -like "*$dataClass*") { 
                $folder = "$PSScriptRoot\TestFiles\NumericPrecisionSearch"
                $filePath = "$folder\$Name"
                $newName = $Name -replace $FTSPre, $AFTName
                Rename-Item -LiteralPath $filePath $newName  

                $count = 1
                $NumOfFiles = 1

                Write-Host "File $count / $NumOfFiles" 
                NetworkShare $newName $folder

                }
            }
    }
Write-Host "
______________________________________
     PHASE II: Run in Background

-- Closing Numeric Presision Search --

$count / $NumOfFiles Files Handled
______________________________________
"
}





# rename files back to original

Get-Childitem -Path "$PSScriptRoot\TestFiles" -File -Filter "*.txt" |
    ForEach-Object { 

        $filePath = $_.FullName
        $oldName = $_.Name -replace $AFTName, $FTSPre

        Rename-Item -LiteralPath $filePath $OldName      

        }

Get-Childitem -Path "$PSScriptRoot\TestFiles\NumericPrecisionSearch" -File -Filter "*.txt" |
    ForEach-Object { 

        $filePath = $_.FullName
        $oldName = $_.Name -replace $AFTName, $FTSPre

        Rename-Item -LiteralPath $filePath $OldName      

        }

# Create Outlook Message and Send

#$csvHTML = $LogPath  | ConvertTo-Html

#$body = @"
#    ....  
#"@ + ($csvHTML[5..($csvHTML.length-2)] | Out-String)

# $Outlook = New-Object -ComObject Outlook.Application
#$Mail = $Outlook.CreateItem(0)
#$Mail.To = "mark.simmons@aig.com"
#$Mail.Subject = "AUTOMATED SEND FROM DLP TEST: $Host"
#$Mail.Body =$body
#$Mail.Attachments = $LogPath
#$Mail.Attachments = "complete path to attachment 2, remove line if not needed."
#$Mail.Attachments = "complete path to attachment 3, remove line if not needed."
#$Mail.Send()  

Write-Host "
*************************
Automated Test is Closing
*************************"