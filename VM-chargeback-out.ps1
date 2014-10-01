#########################################################################################
# Имя: VM-chargeback-out.ps1
# Описание: Выгрузка детализации использования ВМ из MySQL в Excel
#########################################################################################
clear

# Год выгрузки детализации
$count_year=2014

# Выполнение MySQL-команд
function Execute-MySQLQuery([string]$query) { 
  [void][System.Reflection.Assembly]::LoadWithPartialName("MySql.Data")
  $connMySQL = "Server=st-morozov8.office.custis.ru;User=admin;Pwd=Nt[yj>pth5;database=chargeback;"
  $cmd = New-Object MySql.Data.MySqlClient.MySqlCommand($query, $connMySQL)    # Create SQL command
  $dataAdapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($cmd)      # Create data adapter from query command
  $dataSet = New-Object System.Data.DataSet                                    # Create dataset
  $dataAdapter.Fill($dataSet, "data")                                          # Fill dataset from data adapter, with name "data"              
  $cmd.Dispose()
  return $dataSet.Tables["data"]                                               # Returns an array of results
}

# Создание Excel-файла
$excel = New-Object -ComObject Excel.Application
$excel.SheetsInNewWorkbook = 5 # Число листов во вновь созданной книге Excel (по умолчанию 3)
$excel.visible = $true
$workbook = $excel.workbooks.add()
$workbook.author = "Dmitry Morozov"
$workbook.title = "VM detalization"
$workbook.subject = "Demonstrating the Power of PowerShell"

# Переименование листов
$workbook.worksheets.item(1).Name = "INFRA" 
$workbook.worksheets.item(2).Name = "TN" 
$workbook.worksheets.item(3).Name = "GPB" 
$workbook.worksheets.item(4).Name = "NORDEA" 
$workbook.worksheets.item(5).Name = "GS" 

# Заполнения листа детализации для подразделения
function List ($department_name)
{
    # Соотношение название ПП с их идентификаторами в VMware
    switch ($department_name)
    {
        'INFRA' {$dep='IN'}
        'TN' {$dep='TN'}
        'GPB' {$dep='FI'}
        'NORDEA' {$dep='NO'}
        'GS' {$dep='GO'}
        Default {}
    }

    $sheet = $workbook.worksheets.item($department_name) # Выбираем лист книги

    $row = 3    # Первые строки таблицы занята под шапку, поэтому данные будут записаны с третьей строки

    # Задаем стили ячеек шапки
    $lineStyle = "microsoft.office.interop.excel.xlLineStyle" -as [type]
    $colorIndex = "microsoft.office.interop.excel.xlColorIndex" -as [type]
    $borderWeight = "microsoft.office.interop.excel.xlBorderWeight" -as [type]
    $chartType = "microsoft.office.interop.excel.xlChartType" -as [type]
    for($b = 1; $b -le 2; $b++)
    {
        $sheet.cells.item(1,$b).font.bold = $true
        #$sheet.cells.item(1,$b).borders.LineStyle = $lineStyle::xlDashDot
        $sheet.cells.item(1,$b).borders.ColorIndex = $colorIndex::xlColorIndexAutomatic
        $sheet.cells.item(1,$b).borders.weight = $borderWeight::xlMedium
    }
 
    # Дадим осмысленные имена столбцам таблицы
    $sheet.cells.item(1,1) = "Название ВМ"
    for ($d = 1; $d -lt 13; $d++)
    {
        # Год начинается с декабря, функция смещения
        if ($d -eq 1)
        {
            $month=12
            $year=$count_year-1
        }
        else
        {
            $month=$d-1
            $year=$count_year
        }

        # Печатаем шапку
        $sheet.cells.item(1,5*($d-1)+2) = "$month/$year"
        $sheet.cells.item(2,5*($d-1)+2) = "HDD, Гб"
        $sheet.cells.item(2,5*($d-1)+3) = "HDD+backup, Гб"
        $sheet.cells.item(2,5*($d-1)+4) = "RAM, Гб"
        $sheet.cells.item(2,5*($d-1)+5) = "CPU, шт"
        $sheet.cells.item(2,5*($d-1)+6) = "Socets, шт"
    }

    $last_year=$count_year-1
    # Выбор списка ВМ ПП, для которых в БД есть информация об использовании в расчетном году
    $query = "SELECT name, department, timestamp FROM vm WHERE department='$dep' AND ((month(timestamp)<12 AND year(timestamp)='$count_year') OR (month(timestamp)=12 AND year(timestamp)='$last_year')) GROUP BY name ORDER BY name;"
    $result = Execute-MySQLQuery $query
    # Для каждой ВМ выгружаем статистику использования
    for ($i = 1; $i -lt $result.Length; $i++)
    { 
        $name_vm = $result[$i].name
        $sheet.cells.item(($row+$i-1),1) = $name_vm

        # За каждый месяц
        for ($d = 1; $d -lt 13; $d++)
        {
                if ($d -eq 1)
                {
                    $month=12
                    $year=$count_year-1
                }
                else
                {
                    $month=$d-1
                    $year=$count_year
                }
            $query1 = "SELECT name, MAX(cpu) as cpu, round(MAX(mem), 0) as mem, round(MAX(store), 0) as store, department, timestamp FROM vm WHERE name='$name_vm' AND month(timestamp)='$month' AND year(timestamp)='$year' GROUP BY name ORDER BY name;"
            $result1 = Execute-MySQLQuery $query1
        
            # Число сокетов - если 0, то не надо писать
            $sockets = [Math]::Max($result1[1].mem,$result1[1].cpu) 
            if ($sockets -eq 0) {$sockets=""}

            $sheet.cells.item(($row+$i-1),5*($d-1)+2) = $result1[1].store
            $sheet.cells.item(($row+$i-1),5*($d-1)+3) = ($result1[1].store)*2.5
            $sheet.cells.item(($row+$i-1),5*($d-1)+4) = $result1[1].mem
            $sheet.cells.item(($row+$i-1),5*($d-1)+5) = $result1[1].cpu
            $sheet.cells.item(($row+$i-1),5*($d-1)+6) = $sockets
       }
    }

    $range = $sheet.usedRange
    $maxRows = $range.rows.count

    $functions = $excel.WorkSheetfunction

    for ($d = 1; $d -lt 13; $d++)
    {
        #$sheet.range("a${Sumrow}:b$Sumrow").font.bold = "true"
        $first = "AU" + $row
        $last = "AU" + ($result.Length+1)

        $rng = $sheet.Range($first,$last)
        $sheet.cells.item($maxRows+1,5*($d-1)+2) = $functions.sum($rng)
    }
       
    
    #$sheet.cells.item($maxRows+1,5*($d-1)+6) = $functions.sum($rng)
    #$sheet.cells.item($maxRows+1,2) = "Sum $rangeString"
    #$rangeString = $range.address().tostring() -replace "\$",''

    #TODO: Функции суммирования и итоговая стоимость для ПП

    # Выравнивание колонок
    #$range = $sheet.usedRange
    $range.EntireColumn.AutoFit() | Out-Null
}

#List("INFRA")
#List("TN")
#List("GPB")
List("NORDEA")
#List("GS")

# Сохранение Excel-файла
$strPath = "C:\Users\dmorozov\Desktop\VM-detalization.xlsx"
if(Test-Path $strPath)
{
    Remove-Item $strPath
}
$excel.ActiveWorkbook.SaveAs($strPath)

$sheet = $null
$range = $null

# Закрытие Excel 
$workbook.Close($false)
$excel.Quit()
$excel = $null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()