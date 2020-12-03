Function Backup-Folder{

Param(
    [Parameter(Mandatory=$true)]$SrcFilePath,
    [Parameter(Mandatory=$true)]$DstFilePath,
    [Parameter(Mandatory=$False, ParameterSetName="TillDate")][datetime]$TillDate,
    [Parameter(Mandatory=$False, ParameterSetName="FromDate")][datetime]$FromDate,
    [Parameter(Mandatory=$False, ParameterSetName="OnDate")][datetime]$OnDate,
    [Parameter(Mandatory=$False, ParameterSetName="OlderThanDays")][decimal]$OlderThanDays,
    [Parameter(Mandatory=$False, ParameterSetName="LastDays")][decimal]$LastDays
    )
    
    $CurrentDate = Get-Date
    [string[]]$ExcludeExtensionList=@()

    If((Test-Path $SrcFilePath) -eq $false){
        Write-Host "Source path wasn't found." -ForegroundColor Red
        return
    }
    else{
        Write-Host "Source path was found." -ForegroundColor Green
        Switch (Read-Host "Do you want to move the items recursive? `nYes\No (Default: No)"){
            Yes {Write-Host "Recursive enabled" -ForegroundColor Green; $EnableResurcive = $true}
            No {Write-Host "Recursive disabled" -ForegroundColor Red; $EnableResurcive = $false}
            default {Write-Host "Recursive disabled" -ForegroundColor Red; $EnableResurcive = $false}
        }
        $ExcludeExtensionList = Read-Host "Enter extenstions which you want to exclude (e.g. .exe,.csv)"
        $ExcludeExtensionList = $ExcludeExtensionList.Split(',').Split(' ')
    }
    If((Test-Path $DstFilePath) -eq $false){
        Try{
            Write-Host "Destination path wasn't found. Created new directory." -ForegroundColor Green
            New-Item $DstFilePath -ItemType directory
        }
        catch{
            Write-Host "Couldn't create the destination directory. Please check your input."
            return
        }
    }
    else{
        Write-Host "Destination Path exists." -ForegroundColor Green
    }
    $Items = Switch($PSCmdlet.ParameterSetName){
        TillDate {Get-ChildItem $SrcFilePath -Recurse $true}
        FromDate {Write-Host "B"}
        OnDate {Write-Host "C"}
        OlderThanDays {Write-Host "D"}
        LastDays {Write-Host "E"}
        default {"Error. Please check your input."; return}
    }
}