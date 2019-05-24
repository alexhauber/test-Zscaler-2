# Edit This item to change the DropDown Values

[array]$SPI = @("Peoples Republic of China ID","UK NIN ID","UK Sort Code","CC - Visa Number")
$SPIChoice
$checkBox_HTTPSPost_choice = $False
$checkBox_HTTPPost_choice = $False
$checkBox_Print_choice = $False
$checkBox_Network_choice = $False
$checkBox_PrecNum_choice = $False
$checkBox_AutoSelectNetFolder_choice = $True
$executeClicked = $False
$networkFolder = "init"
#[array]$EventType = @("HTTPS Post", "Print", "Create File")
Add-Type -AssemblyName PresentationFramework

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
[System.Windows.Forms.Application]::EnableVisualStyles()


function Ask-Folder {

    $browse = New-Object System.Windows.Forms.FolderBrowserDialog
    $browse.SelectedPath = "C:\"
    $browse.ShowNewFolderButton = $true
    $browse.Description = "Select a Network Directory"

    $loop = $true
    while($loop)
    {
        if ($browse.ShowDialog() -eq "OK")
        {
        $loop = $false
		
		#Insert your script here
		
        } else
        {
            $res = [System.Windows.Forms.MessageBox]::Show("Clicking Cancel will revert to the Home Drive auto-detection option. 
            
Would you like to try again or Cancel?", "Select a location", [System.Windows.Forms.MessageBoxButtons]::RetryCancel)
            if($res -eq "Cancel")
            {
                #Ends script
                return
            }
        }
    }

    $browse.SelectedPath
    $global:networkFolder = $browse.SelectedPath
    $browse.Dispose()
} 


# This Function Returns the Selected Value and Closes the Form

function Execute_Button {

    $global:executeClicked = $True
    
    #get network drive window process
    if ($checkBox_Network.Checked -OR $checkBox_PrecNum.Checked)         {
        $checkBox_Network.Checked = $True
        $script:checkBox_Network_choice = $True

        if ($FolderSelect_Button_UserSelect.Checked){
            $checkBox_AutoSelectNetFolder_choice = $False 
            Ask-Folder
            write-host "
NETWORK FOLDER SELECTED:
$networkFolder
"
            if ($networkFolder -eq "init") {
                $global:executeClicked = $False

        write-host "
NETWORK FOLDER SELECTED 2:
$networkFolder
"
    
                }
            }
            
            else { 
                $checkBox_AutoSelectNetFolder_choice = $True
                #attempts home-drive path auto-detection
                Get-PSDrive -PSProvider FileSystem | ForEach-Object { 
        
                    $driveRoot = $_.DisplayRoot
                    if ($driveRoot -like "*$userName*") { $HomeDrivePath = $driveRoot}

                    }
                #below comment out for home drive auto-detection failed message
                #$HomeDrivePath = ""

                if ($HomeDrivePath.Length -eq 0) { 
                    $global:networkFolder = "init"
                    [System.Windows.MessageBox]::Show('Auto-detection of Home Drive failed. 
Returning to main window.

"Network Share" option will revert to "Select drive manually". ')
                    $FolderSelect_Button_UserSelect.Checked
                    $global:executeClicked = $False
                    
                    }

                else { $global:networkFolder = $driveRoot } 
            }
                
       }
       

        $global:SPIChoice = $DropDownSPI.SelectedItem.ToString()

        if ($checkBox_HTTPSPost.Checked)     {  $script:checkBox_HTTPSPost_choice = $True }

        if ($checkBox_HTTPPost.Checked)     {  $script:checkBox_HTTPPost_choice = $True }

        if ($checkBox_Print.Checked)         {  $script:checkBox_Print_choice = $True }

        if ($checkBox_PrecNum.Checked)         {  $script:checkBox_PrecNum_choice = $True 
                                                  $script:checkBox_Network_choice = $True                                             
                                                }

        if ($FolderSelect_Button_UserSelect.Checked) { $checkBox_AutoSelectNetFolder_choice = $False } 


        if ($networkFolder -eq "init" -and $checkBox_Network.Checked -eq $True) {
            $global:executeClicked = $False
            $FolderSelect_Button_HomeDrive.Checked = $True
            #$Form.Close()
            #Load-MainForm
        }
        else {
            $Form.Close()
  
            }

}

#function Load-MainForm {

## init form
$Form = New-Object System.Windows.Forms.Form

$Form.width = 458
$Form.height = 550
$Form.Text = ”DLP Automated Test”

$fe_Start_x = 4
$fe_Start_y = 4
$sep_x = 12
$col1_x = 14
$col2_x = $col1_x + 17

$instructions_Width = 390
$instrutions_height = 24
$sep_w = 415

$SPIInstrucitons = new-object System.Windows.Forms.Label
$SPIInstrucitons.Location = new-object System.Drawing.Size($col1_x,($fe_Start_y+8))
$SPIInstrucitons.size = new-object System.Drawing.Size($instructions_Width,$instrutions_height)
$SPIInstrucitons.Text = "1. Select desired SPI to test. To test all, select ""-- ALL SPI --""."
$Form.Controls.Add($SPIInstrucitons)

$SPI_y = 49
## init SPI dropdown
$DropDownSPI = new-object System.Windows.Forms.ComboBox
$DropDownSPI.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$DropDownSPI.Sorted = $true
$DropDownSPI.Location = new-object System.Drawing.Size(($col2_x + 40),$SPI_y)
$DropDownSPI.Size = new-object System.Drawing.Size(325,30)
#$DropDownSPI.AutoCompleteMode = $true
#$DropDownSPI.CausesValidation = $true

$DropDownSPI.Items.Add("-- ALL SPI --")

##This is the current version which uses the premade text files
Get-Childitem -Path "$PSScriptRoot\TestFiles" -File -Filter "*.txt" |
    ForEach-Object { 
        
        #Write-Host "File $count / $NumOfFiles" 
        [void] $DropDownSPI.Items.Add($_)
        
        }

##This will be turned on when the test arrays are biult into the ps1 and test files are created dynamically
#ForEach ($Item in $SPI) {
#[void] $DropDownSPI.Items.Add($Item)
#}

$Form.Controls.Add($DropDownSPI)

$DropDownSPILabel = new-object System.Windows.Forms.Label
$DropDownSPILabel.Location = new-object System.Drawing.Size($col2_x,$SPI_y)
$DropDownSPILabel.size = new-object System.Drawing.Size(100,20)
$DropDownSPILabel.Text = "SPI"
$DropDownSPI.SelectedItem = $DropDownSPI.Items[0]
$Form.Controls.Add($DropDownSPILabel)

$SPISep_y = 90
$SPISep = new-object System.Windows.Forms.Label
$SPISep.Text = ""
$SPISep.BorderStyle = 'Fixed3D'
$SPISep.Location = new-object System.Drawing.Size($sep_x,$SPISep_y)
$SPISep.size = new-object System.Drawing.Size($sep_w,3)
$Form.Controls.Add($SPISep)

$EI_y = $SPISep_y + 12
$EventInstrucitons = new-object System.Windows.Forms.Label
$EventInstrucitons.Location = new-object System.Drawing.Size($col1_x,$EI_y)
$EventInstrucitons.size = new-object System.Drawing.Size($instructions_Width,$instrutions_height)
$EventInstrucitons.Text = "2. Select desired Event Type(s) to test."
$Form.Controls.Add($EventInstrucitons)


## init event type checkboxes
 
$checkBox_Start_y = $EI_y + 30

$checkBox_HTTPSPost = New-Object System.Windows.Forms.CheckBox
$checkBox_HTTPSPost.UseVisualStyleBackColor = $True
$System_Drawing_Size_ch = New-Object System.Drawing.Size
$System_Drawing_Size_ch.Width = 124
$System_Drawing_Size_ch.Height = 24
$checkBox_HTTPSPost.Size = $System_Drawing_Size_ch
$checkBox_HTTPSPost.TabIndex = 2
$checkBox_HTTPSPost.Text = "HTTPS Post"
$System_Drawing_Point_ch = New-Object System.Drawing.Point
$System_Drawing_Point_ch.X = $col2_x
$System_Drawing_Point_ch.Y = $checkBox_Start_y
$checkBox_HTTPSPost.Location = $System_Drawing_Point_ch
$checkBox_HTTPSPost.DataBindings.DefaultDataSourceUpdateMode = 1
$checkBox_HTTPSPost.Name = "checkBox_HTTPSPost"
$checkBox_HTTPSPost.Checked = $True
$Form.Controls.Add($checkBox_HTTPSPost)

$checkBox_HTTPPost = New-Object System.Windows.Forms.CheckBox
$checkBox_HTTPPost.UseVisualStyleBackColor = $True
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 124
$System_Drawing_Size.Height = 24
$checkBox_HTTPPost.Size = $System_Drawing_Size
$checkBox_HTTPPost.TabIndex = 2
$checkBox_HTTPPost.Text = "HTTP Post"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = $col2_x
$System_Drawing_Point.Y = $checkBox_Start_y + 29
$checkBox_HTTPPost.Location = $System_Drawing_Point
$checkBox_HTTPPost.DataBindings.DefaultDataSourceUpdateMode = 1
$checkBox_HTTPPost.Name = "checkBox_HTTPPost"
$checkBox_HTTPPost.Checked = $False
$Form.Controls.Add($checkBox_HTTPPost)

$checkBox_Print = New-Object System.Windows.Forms.CheckBox
$checkBox_Print.UseVisualStyleBackColor = $True
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 124
$System_Drawing_Size.Height = 24
$checkBox_Print.Size = $System_Drawing_Size
$checkBox_Print.TabIndex = 2
$checkBox_Print.Text = "Print"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = $col2_x
$System_Drawing_Point.Y = $checkBox_Start_y + 59
$checkBox_Print.Location = $System_Drawing_Point
$checkBox_Print.DataBindings.DefaultDataSourceUpdateMode = 1
$checkBox_Print.Name = "checkBox_Print"
$checkBox_Print.Checked = $True
$Form.Controls.Add($checkBox_Print)

$checkBox_Network = New-Object System.Windows.Forms.CheckBox
$checkBox_Network.UseVisualStyleBackColor = $True
$System_Drawing_Size_cn = New-Object System.Drawing.Size
$System_Drawing_Size_cn.Width = 124
$System_Drawing_Size_cn.Height = 24
$checkBox_Network.Size = $System_Drawing_Size_cn
$checkBox_Network.TabIndex = 2
$checkBox_Network.Text = "Network Share"
$System_Drawing_Point_cn = New-Object System.Drawing.Point
$System_Drawing_Point_cn.X = $col2_x + 180
$System_Drawing_Point_cn.Y = $checkBox_Start_y
$checkBox_Network.Location = $System_Drawing_Point_cn
$checkBox_Network.DataBindings.DefaultDataSourceUpdateMode = 1
$checkBox_Network.Name = "checkBox_Network"
$checkBox_Network.Checked = $True
$Form.Controls.Add($checkBox_Network)

# Network Share Folder Option
$groupBox_FolderSelect = New-Object System.Windows.Forms.GroupBox #create the group box
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 205
$System_Drawing_Size.Height = 80
$groupBox_FolderSelect.size = $System_Drawing_Size
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = $col2_x + 186
$System_Drawing_Point.Y = $checkBox_Start_y + 10
$groupBox_FolderSelect.Location = $System_Drawing_Point #location of the group box (px) in relation to the primary window's edges (length, height)
$groupBox_FolderSelect.text = "Drive Location:" #labeling the box
$Form.Controls.Add($groupBox_FolderSelect) #activate the group box
$FolderSelect_Button_HomeDrive = New-Object System.Windows.Forms.RadioButton #create the radio button
$FolderSelect_Button_HomeDrive.Location = new-object System.Drawing.Point(15,15) #location of the radio button(px) in relation to the group box's edges (length, height)
$FolderSelect_Button_HomeDrive.size = New-Object System.Drawing.Size(180,24) #the size in px of the radio button (length, height)
$FolderSelect_Button_HomeDrive.Checked = $true #is checked by default
$FolderSelect_Button_HomeDrive.Text = "Auto-detect Home Drive" #labeling the radio button
$groupBox_FolderSelect.Controls.Add($FolderSelect_Button_HomeDrive) #activate the inside the group box
$FolderSelect_Button_UserSelect = New-Object System.Windows.Forms.RadioButton #create the radio button
$FolderSelect_Button_UserSelect.Location = new-object System.Drawing.Point(15,43.5) #location of the radio button(px) in relation to the group box's edges (length, height)
$FolderSelect_Button_UserSelect.size = New-Object System.Drawing.Size(180,24) #the size in px of the radio button (length, height)
$FolderSelect_Button_UserSelect.Checked = $false #is checked false by default
$FolderSelect_Button_UserSelect.Text = "Select drive manually" #labeling the radio button
$groupBox_FolderSelect.Controls.Add($FolderSelect_Button_UserSelect) #activate the inside the group box


$ExeSep = new-object System.Windows.Forms.Label
$ExeSep.Text = ""
$ExeSep.BorderStyle = 'Fixed3D'
$ExeSep.Location = new-object System.Drawing.Size($sep_x,235)
$ExeSep.size = new-object System.Drawing.Size($sep_w,3)
$Form.Controls.Add($ExeSep)

#$ExecuteInstrucitons_height = $instrutions_height + 60
$ExecuteInstrucitons = new-object System.Windows.Forms.Label
$ExecuteInstrucitons.Location = new-object System.Drawing.Size($col1_x,253)
$ExecuteInstrucitons.size = new-object System.Drawing.Size($instructions_Width,$instrutions_height)
$ExecuteInstrucitons.Text = "3. Select Precision Search Type(s). "
$Form.Controls.Add($ExecuteInstrucitons)
$ExecuteInstrucitons = new-object System.Windows.Forms.Label
$ExecuteInstrucitons.Location = new-object System.Drawing.Size($col2_x,277)
$ExecuteInstrucitons.size = new-object System.Drawing.Size($instructions_Width,$instrutions_height)
$ExecuteInstrucitons.Text = "For Network Share only."
$Form.Controls.Add($ExecuteInstrucitons)
$ExecuteInstrucitons = new-object System.Windows.Forms.Label
$ExecuteInstrucitons.Location = new-object System.Drawing.Size($col2_x,302)
$ExecuteInstrucitons.size = new-object System.Drawing.Size($instructions_Width,$instrutions_height)
$ExecuteInstrucitons.Text = "Note:"
$Form.Controls.Add($ExecuteInstrucitons)
$ExecuteInstrucitons_height = $instrutions_height + 10
$ExecuteInstrucitons = new-object System.Windows.Forms.Label
$ExecuteInstrucitons.Location = new-object System.Drawing.Size($col2_x,327.8)
$ExecuteInstrucitons.size = new-object System.Drawing.Size($instructions_Width,$ExecuteInstrucitons_height)
$ExecuteInstrucitons.Text = "This will cause a standard Network Share test on the selected SPI regardless of what is selected in ""2."" above."
$Form.Controls.Add($ExecuteInstrucitons)
#$ExecuteInstrucitons = new-object System.Windows.Forms.Label
#$ExecuteInstrucitons.Location = new-object System.Drawing.Size($col1_x,208)
#$ExecuteInstrucitons.size = new-object System.Drawing.Size($instructions_Width,$instrutions_height)
#$ExecuteInstrucitons.Text = "3. Select Precision Search Type(s).\r\n For Network Share only.Note: This will cause a standard network share test, regardless of what is selected in ""2."" above."
#$Form.Controls.Add($ExecuteInstrucitons)

$checkBox_PrecNum = New-Object System.Windows.Forms.CheckBox
$checkBox_PrecNum.UseVisualStyleBackColor = $True
$System_Drawing_Size_pn = New-Object System.Drawing.Size
$System_Drawing_Size_pn.Width = 124
$System_Drawing_Size_pn.Height = 24
$checkBox_PrecNum.Size = $System_Drawing_Size_pn
$checkBox_PrecNum.TabIndex = 2
$checkBox_PrecNum.Text = "Numeric"
$System_Drawing_Point_pn = New-Object System.Drawing.Point
$System_Drawing_Point_pn.X = $col2_x
$System_Drawing_Point_pn.Y = $checkBox_Start_y + 240
$checkBox_PrecNum.Location = $System_Drawing_Point_pn
$checkBox_PrecNum.DataBindings.DefaultDataSourceUpdateMode = 1
$checkBox_PrecNum.Name = "checkBox_PrecNum"
$checkBox_PrecNum.Checked = $True
$Form.Controls.Add($checkBox_PrecNum)

$ExeSep = new-object System.Windows.Forms.Label
$ExeSep.Text = ""
$ExeSep.BorderStyle = 'Fixed3D'
$ExeSep.Location = new-object System.Drawing.Size($sep_x,405)
$ExeSep.size = new-object System.Drawing.Size($sep_w,3)
$Form.Controls.Add($ExeSep)

$ExecuteInstrucitons = new-object System.Windows.Forms.Label
$ExecuteInstrucitons.Location = new-object System.Drawing.Size($col1_x,418)
$ExecuteInstrucitons.size = new-object System.Drawing.Size($instructions_Width,$instrutions_height)
$ExecuteInstrucitons.Text = "4. Click the Execute button."
$Form.Controls.Add($ExecuteInstrucitons)

## init execut button
$Button_Execute = new-object System.Windows.Forms.Button
$Button_Execute.Location = new-object System.Drawing.Size($col2_x,450)
$Button_Execute.Size = new-object System.Drawing.Size(370,35)
$Button_Execute.Text = "Execute"
$Button_Execute.Add_Click({Execute_Button})
$form.Controls.Add($Button_Execute)

## activate form
$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()
## -- 

if ($executeClicked -eq $True) {


    $checkBox_HTTPSPost_choice_pass = "$"+$checkBox_HTTPSPost_choice
    $checkBox_HTTPPost_choice_pass = "$"+$checkBox_HTTPPost_choice
    $checkBox_Print_choice_pass = "$"+$checkBox_Print_choice
    $checkBox_Network_choice_pass = "$"+$checkBox_Network_choice
    $checkBox_PrecNum_choice_pass = "$" + $checkBox_PrecNum_choice
    $checkBox_AutoSelectNetFolder_choice_pass = "$" + $checkBox_AutoSelectNetFolder_choice
    $ScriptFolder = split-path -parent $MyInvocation.MyCommand.Definition
    $ScriptPath = "$ScriptFolder\Process.ps1"
    $cmd = "&’$ScriptPath' '$SPIChoice' $checkBox_HTTPSPost_choice_pass $checkBox_HTTPPost_choice_pass $checkBox_Print_choice_pass $checkBox_Network_choice_pass $checkBox_PrecNum_choice_pass $networkFolder"

    Invoke-Expression $cmd
    }


#}

#Load-MainForm
  