Import-Module ImportExcel

# Output folder
$OutputFolder = "C:\Users\a.gadallah\OneDrive - Al Jasser Holding\Desktop\May\Oracle"

if (!(Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

# Excel files
$BlockVolumeExcel = Join-Path $OutputFolder "BlockVolumeBackups.xlsx"
$BootVolumeExcel  = Join-Path $OutputFolder "BootVolumeBackups.xlsx"

# Remove old files if exist
if (Test-Path $BlockVolumeExcel) {
    Remove-Item $BlockVolumeExcel -Force
}

if (Test-Path $BootVolumeExcel) {
    Remove-Item $BootVolumeExcel -Force
}

Write-Host "Getting compartments..."

# Get tenancy OCID from config
$ConfigFile = "$HOME\.oci\config"

$TenancyId = Select-String -Path $ConfigFile -Pattern "^tenancy=" |
ForEach-Object { $_.Line.Split("=")[1].Trim() }

Write-Host "Using Tenancy: $TenancyId"

# Get all compartments
$CompartmentsJson = oci iam compartment list `
    --compartment-id $TenancyId `
    --compartment-id-in-subtree true `
    --all `
    --output json

$Compartments = ($CompartmentsJson | ConvertFrom-Json).data

# Add root tenancy manually
$RootCompartment = [PSCustomObject]@{
    id = $TenancyId
    name = "Root"
}

$Compartments += $RootCompartment

# =============================================
# BLOCK VOLUME BACKUPS
# =============================================

Write-Host "Collecting Block Volume Backups..."

$BlockResults = @()

foreach ($Comp in $Compartments) {

    Write-Host "Checking Block Volume Backups in compartment: $($Comp.name)"

    try {

        $BlockBackupsJson = oci bv backup list `
            --compartment-id $Comp.id `
            --all `
            --output json

        $BlockBackups = ($BlockBackupsJson | ConvertFrom-Json).data

        foreach ($Backup in $BlockBackups) {

            $MonthSheet = (Get-Date $Backup."time-created").ToString("yyyy-MM")

            $BlockResults += [PSCustomObject]@{
                Month            = $MonthSheet
                CompartmentName  = $Comp.name
                BackupName       = $Backup."display-name"
                BackupId         = $Backup.id
                VolumeId         = $Backup."volume-id"
                SizeInGB         = $Backup."size-in-gbs"
                BackupType       = $Backup.type
                LifecycleState   = $Backup."lifecycle-state"
                TimeCreated      = $Backup."time-created"
                ExpirationTime   = $Backup."expiration-time"
                UniqueSizeInGB   = $Backup."unique-size-in-gbs"
            }
        }
    }
    catch {
        Write-Host "Error reading compartment $($Comp.name)"
    }
}

# Export Block Volume Backups Monthly Sheets
$BlockMonths = $BlockResults.Month | Sort-Object -Unique

foreach ($Month in $BlockMonths) {

    Write-Host "Exporting Block Volume month: $Month"

    $MonthData = $BlockResults | Where-Object { $_.Month -eq $Month }

    $MonthData | Export-Excel `
        -Path $BlockVolumeExcel `
        -WorksheetName $Month `
        -AutoSize `
        -FreezeTopRow `
        -BoldTopRow
}

# =============================================
# BOOT VOLUME BACKUPS
# =============================================

Write-Host "Collecting Boot Volume Backups..."

$BootResults = @()

foreach ($Comp in $Compartments) {

    Write-Host "Checking Boot Volume Backups in compartment: $($Comp.name)"

    try {

        $BootBackupsJson = oci bv boot-volume-backup list `
            --compartment-id $Comp.id `
            --all `
            --output json

        $BootBackups = ($BootBackupsJson | ConvertFrom-Json).data

        foreach ($Backup in $BootBackups) {

            $MonthSheet = (Get-Date $Backup."time-created").ToString("yyyy-MM")

            $BootResults += [PSCustomObject]@{
                Month            = $MonthSheet
                CompartmentName  = $Comp.name
                BackupName       = $Backup."display-name"
                BackupId         = $Backup.id
                BootVolumeId     = $Backup."boot-volume-id"
                SizeInGB         = $Backup."size-in-gbs"
                BackupType       = $Backup.type
                LifecycleState   = $Backup."lifecycle-state"
                TimeCreated      = $Backup."time-created"
                ExpirationTime   = $Backup."expiration-time"
                UniqueSizeInGB   = $Backup."unique-size-in-gbs"
            }
        }
    }
    catch {
        Write-Host "Error reading compartment $($Comp.name)"
    }
}

# Export Boot Volume Backups Monthly Sheets
$BootMonths = $BootResults.Month | Sort-Object -Unique

foreach ($Month in $BootMonths) {

    Write-Host "Exporting Boot Volume month: $Month"

    $MonthData = $BootResults | Where-Object { $_.Month -eq $Month }

    $MonthData | Export-Excel `
        -Path $BootVolumeExcel `
        -WorksheetName $Month `
        -AutoSize `
        -FreezeTopRow `
        -BoldTopRow
}

Write-Host ""
Write-Host "Export Completed Successfully"
Write-Host ""
Write-Host "Block Volume File: $BlockVolumeExcel"
Write-Host "Boot Volume File : $BootVolumeExcel"