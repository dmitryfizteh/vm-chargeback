#########################################################################################
# Имя: VM-chargeback-in.ps1
# Описание: Загрузка детализации использования ВМ в MySQL, вывод проблемных ВМ
# и краткого состояния виртуальной среды
# Автор: Dmitry Morozov
# Версия: 1.0
#########################################################################################

cls

# Определения и коннекты
#Add-pssnapin VMware*

$vc1 = "vmcl-vcenter"
$vc2 = "vm-vcenter-02"
$vc3 = "vm-vcenter03"

#TODO: автоматический ввод даты (с возможностью ручного выбора)
$timestamp = "2014-09-30"
#$timestamp = Read-Host "На какую дату вводятся данные? (YYYY-MM-DD)"

# Суммы ресурсов всех кластеров, используемых ВМ
$all_vm_cpu=0
$all_vm_mem=0
$all_vm_store_Tb=0

# Суммы ресурсов всех кластеров
$all_cl_cpu=0
$all_cl_mem=0
$all_cl_store_Tb=0

# Выполнение MySQL-команд
function Execute-MySQLQuery([string]$query) { 
  [void][System.Reflection.Assembly]::LoadWithPartialName("MySql.Data")
  #TODO: убрать реквизиты MySQL-сервера в определения
  $connMySQL = "Server=st-morozov8.office.custis.ru;User=admin;Pwd=Nt[yj>pth5;database=chargeback;"
  $cmd = New-Object MySql.Data.MySqlClient.MySqlCommand($query, $connMySQL)    # Create SQL command
  $dataAdapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($cmd)      # Create data adapter from query command
  $dataSet = New-Object System.Data.DataSet                                    # Create dataset
  $dataAdapter.Fill($dataSet, "data")                                          # Fill dataset from data adapter, with name "data"              
  $cmd.Dispose()
  return $dataSet.Tables["data"]                                               # Returns an array of results
}

# Функция загрузки детализации для каждого кластера
function Claster_discover ($claster_name)
{
    Write-Host "Обрабатываем кластер $claster_name" -ForegroundColor "Green"
    Connect-VIServer $claster_name | Out-Null

    # Суммарные ресурсы CPU и RAM кластера
    $sum_cl_cpu=0
    $sum_cl_mem=0

    # Сбор информации о ресурсах кластера
    Get-VMHost | Select Name,Model,NumCpu,MemoryTotalGB,CpuTotalMhz | ft -AutoSize
    $hosts=Get-VMHost | Select Name,NumCpu | Measure-Object -Property NumCpu -Sum
    $sum_cl_cpu=$hosts.Sum
    $hosts=Get-VMHost | Select Name,MemoryTotalGB |Measure-Object -Property MemoryTotalGB -Sum
    $sum_cl_mem=$hosts.Sum
    #TODO: сделать сбор информации по HDD кластера
    #TODO: сделать процент использования CPU кластера как CpuUsageMhz/CpuTotalMhz

    #TODO: сделать сбор информации о ресурсах темплейтов и помещение их в ВМ

    # Сбор информации о ресурсах ВМ кластера
    $vms = Get-VM | Sort-Object Name #| Format-Table Name, ResourcePool, NumCpu, MemoryGB, ProvisionedSpaceGB, Notes
    $without_notes="" # Список ВМ без VMware tools
    # Суммарные ресурсы кластера, используемые ВМ 
    $sum_cpu=0
    $sum_mem=0
    $sum_store=0

    # Добавление информации о ВМ в БД MySQL
    foreach ($item in $vms)
    {
        $name=$item.Name
        $cpu=$item.NumCpu
        $sum_cpu=$sum_cpu+$cpu
        $mem=$item.MemoryGB
        $sum_mem=$sum_mem+$mem
        $store=$item.ProvisionedSpaceGB
        $sum_store=$sum_store+$store

        # Если notes не пустые
        if (($item.Notes).Length -gt 2)
        {   
            # Поле $department - первые 2 буквы названия подразделения  
            $department=($item.Notes).Remove(2)
            # Если department существует
            switch ($department)
            {
                {$_ -in 'IN','FI','TN','GO','NO'} 
                    {
                        $query = "INSERT INTO vm (name, cpu, mem, store, department, timestamp) VALUES ('$name', '$cpu', '$mem', '$store', '$department', '$timestamp');"
                        #$query
                        #$result = Execute-MySQLQuery $query

                        #TODO: Добавить инфо про owner, bug, backup, creation_date, description, OS
                    }
                Default {$without_notes+=($item.Name) +"`n"}
            }      
        }
        else
        {
            $without_notes+=($item.Name) +"`n"
        }
   
    }

    # Вывод ВМ без описаний (невозможно определить подразделение-владельца)
    Write-Host "`nСледующие ВМ без Notes:" -ForegroundColor "Red"
    foreach ($item in $without_notes)  {$item}
    #TODO: перенести в VM-tools.ps1

    # Суммирование потребляемых ресурсов
    $sum_store_Tb = [Math]::Ceiling($sum_store/1024)
    $sum_mem = [Math]::Ceiling($sum_mem)
    $sum_cl_mem = [Math]::Ceiling($sum_cl_mem)
    $cpu_ratio = [Math]::Ceiling($sum_cpu/$sum_cl_cpu*100)
    $mem_ratio = [Math]::Ceiling($sum_mem/$sum_cl_mem*100)
    #TODO: Сделать % использования HDD
    Write-Host "Summary:" -ForegroundColor "Green"
    Write-Host "Всего в кластере $sum_cl_cpu cpu, $sum_cl_mem Gb RAM, xxx Tb HDD."
    Write-Host "Из них ВМ используют $sum_cpu cpu, $sum_mem Gb RAM, $sum_store_Tb Tb HDD."
    Write-Host "Итого, на кластере $cpu_ratio % cpu, $mem_ratio % RAM, xxx % HDD (при максимально возможных 70%)."

    #$all_vm_cpu=$all_vm_cpu+$sum_cpu
    #$all_vm_mem=$all_vm_mem+$sum_mem
    #$all_vm_store_Tb=$all_vm_store_Tb+$sum_store_Tb
    #$all_cl_cpu=$all_cl_cpu+$sum_cl_cpu
    #$all_cl_mem=$all_cl_mem+$sum_cl_mem
    #$all_cl_store_Tb=$all_cl_store_Tb+$sum_cl_store_Tb

    Write-Host "`nКластер $claster_name обработан." -ForegroundColor "Green"
}

#Claster_discover ($vc1)
Claster_discover ($vc2)
Claster_discover ($vc3)

#TODO: Сделать анализ потребляемых ресурсов для всех кластеров
#Write-Host "Всего в кластере $all_cl_cpu cpu, $all_cl_mem Gb RAM, xxx Tb HDD."
#Write-Host "Из них ВМ используют  $all_vm_cpu cpu, $all_vm_mem Gb RAM, $all_vm_store_Tb Tb HDD."
#Write-Host "`nСкрипт завершил свою работу." -ForegroundColor "Green"
