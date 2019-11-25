 #(DSS)

Add-Type -AssemblyName PresentationFramework

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$form = New-Object System.Windows.Forms.Form
$form.Text = "Hello Drop Crew"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(300,200)
$form.FormBorderStyle = "FixedDialog"

$lable1 = New-Object System.Windows.Forms.Label
$lable1.Location = New-Object System.Drawing.Size(8,25)
$lable1.Size = New-Object System.Drawing.Size(275,32)
$lable1.TextAlign = "MiddleCenter"
$lable1.Text = "Where are you located?"
$form.Controls.Add($lable1)

$buckysButton = New-Object System.Windows.Forms.Button
$buckysButton.Text = "Bucky's"
$buckysButton.Location = New-Object System.Drawing.Size(60,80)
$buckysButton.Size = New-Object System.Drawing.Size(100,32)
$buckysButton.Height = 25
$buckysButton.Width = 75
$buckysButton.add_click({$Global:Ans = 1; $form.Close()})
$form.controls.Add($buckysButton)

$yavaButton = New-Object System.Windows.Forms.Button
$yavaButton.Text = "Yavapai"
$yavaButton.Location = New-Object System.Drawing.Size(160,80)
$yavaButton.Size = New-Object System.Drawing.Size(100,32)
$yavaButton.Height = 25
$yavaButton.Width = 75
$yavaButton.add_click({$Global:Ans = 2; $form.Close()})
$form.controls.Add($yavaButton)

#----------------------------------------------------------------------------------------------------------------------------------
$form.ShowDialog()

#The user will be prompted with a location option. There choice
#will choose the correct path of the jisseki file.

$date = Get-Date -UFormat "%m-%d-%y@%H-%M-%S"

Function GetPropertyLocation
{
    param($ButtonAns)

    if ($ButtonAns -eq 1)
    {
        $path = "PATH FOR CASINO 1"
        return $path
    }
    elseif($ButtonAns -eq 2)
    {
        $path = "PATH FOR CASINO 2"
        return $path
    }
    else
    {
        break
    }
}

#Gets the correctly selected location path for the Jisseki file
$jissekiPath = GetPropertyLocation($Ans)

#----------------------------------------------------------------------------------------------------------------------------------
<#    
# Ticket conversion
$latestTicket = Get-ChildItem "PATH FOR TICKETS" | 
Sort-Object CreationTime -Descending | 
Select-Object -First 1 #>

# Path for YC latest file
$YCLatestCurrencyFile = Get-ChildItem "PATH CASINO 2 CURRENCY" | 
Sort-Object CreationTime -Descending | 
Select-Object -First 1

# Path for Bucky's latest file
$BCLatestCurrencyFile = Get-ChildItem "PATH CASINO 1 CURRENCY" | 
Sort-Object CreationTime -Descending | 
Select-Object -First 1

#Heavy lifter----------------------------------------------------------------------------------------------------------------------
#checks if the file exists, if it doesn't then an error will come up, else it will continue with the converting process

if (![System.IO.File]::Exists($jissekiPath))
{
    [System.Windows.MessageBox]::Show("This file doesn't exist", "Error", "Ok", "Error")
}
else
{
    Start-Transcript -Path "PATH FOR TRANSCRIPT BACKUP" -NoClobber
    # Makes a backup of the JISSEKI file
    Copy-Item -Path $jissekiPath -destination PATH FOR JISSEKI BACKUP-$date.CSV

    # This line will import csv file and create the Machine and Box header, 
    #It will also nigate the first column with the date and time
    $finalJissekiFile = Import-CSV $jissekiPath -Header DeleteThis, Machine, Box | Select "Machine", "Box"
    
    #Buckys ***********************************************************************************
    if ($Ans -eq 1)
    {
        $CurrencyPath = "CASINO 1 CURRENCY PATH"
        $DestinationPath = "CASINO 1 CURRENCY BACKUP PATH-$date.CSV"

        # Makes a backup of the JISSEKI file
        Copy-Item -Path $jissekiPath -destination BACKUP PATH-$date.CSV

        #Makes copy of the currency file before any manipulation
        Copy-Item -Path $CurrencyPath -Destination $DestinationPath

        ForEach ($line in $finalJissekiFile)
        {
            $content = Get-Content "LATEST CURRENCY FILE PATH"
            ForEach-Object {$content -replace $line.Box, $line.Machine } |
            Set-Content "LATEST CURRENCY FILE PATH"
            Write-Host "$line"
        }
        #ticketConvert
    }

    # Yavapai *********************************************************************************
    elseif ($Ans -eq 2)
    {
        $CurrencyPath = "CASINO 2 CURRENCY PATH"
        $DestinationPath = "CASINO 2 CURRENCY BACKUP PATH-$date.CSV"

        # Makes a backup of the JISSEKI file
        Copy-Item -Path $jissekiPath -destination BACKUP PATH-$date.CSV

        Copy-Item -Path $CurrencyPath -Destination $DestinationPath

        ForEach ($line in $finalJissekiFile)
        {
            $content = Get-Content "LATEST CURRENCY FILE PATH"
            ForEach-Object {$content -replace $line.Box, $line.Machine } |
            Set-Content LATEST CURRENCY FILE PATH"
            Write-Host "$line"
        }
        #ticketConvert
    }
    
    [System.Windows.MessageBox]::Show("Successful :)")
    Stop-Transcript
}