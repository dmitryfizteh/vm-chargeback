#########################################################################################
# Имя: VM-tools.ps1
# Описание: Выводит списки ВМ с неактуальными VMware Tools, выключенных ВМ
# Автор: Dmitry Morozov
# Версия: 1.0
#########################################################################################

$vc1 = "vmcl-vcenter"
$vc2 = "vm-vcenter-02"
$vc3 = "vm-vcenter03"

cls

# Выводит список ВМ с неактуальными VMware Tools для каждого кластера
function Print_Bad_VMtools ()
{
    Write-Host "`nСписок машин с неактуальными VMwareTools:" -ForegroundColor "Gray"
    get-vm | % { get-view $_.ID } | select Name, PowerState, { Name="hostName"; Expression={$_.guest.hostName}}, @{ Name="ToolsStatus"; Expression={$_.guest.toolsstatus}}, @{ Name="ToolsVersion"; Expression={$_.config.tools.toolsVersion}} | Where-Object {$_.ToolsStatus -ne "toolsOk"} | sort-object name | ft name, ToolsStatus 

    #TODO: Если ВМ в состоянии PowerOff и toolsNotRunning, то не выводить
}

# Вывод выключенных ВМ (не нужно ли их удалить?)
function Print_PowerOff_VMs ()
{
    Write-Host "`nСписок выключенных ВМ:" -ForegroundColor "Gray"
    get-vm | where {$_.PowerState -eq "PoweredOff"} |Format-Table Name
}

# Вывод списка неудаленных снепшотов
function Print_Old_snapshots ()
{
    #Write-Host "`nСписок ВМ со снепшотами:" -ForegroundColor "Gray"
    #Get-VM | Get-Snapshot

    #TODO: Отправить по e-mail напоминание владельцам ВМ
}

# Поиск проблем на каждом кластере
function ClasterFindProblems ($vcenter)
{
    Write-Host "Обрабатываем кластер $vcenter" -ForegroundColor "Green"
    Connect-VIServer -Server $vcenter | Out-Null
    Print_Bad_VMtools
    Print_PowerOff_VMs 
    #TODO: Добавить список неудаленных снепшотов
    #Print_Old_snapshots
}

ClasterFindProblems ($vc1) 
ClasterFindProblems ($vc2)
ClasterFindProblems ($vc3)

#TODO: Отправить списки по e-mail