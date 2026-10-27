Write-Host "Iniciando Excel..."

$excel = $null
$workbook = $null

try {
    $arquivo = "D:\"**NOME_DA_SUA_PASTA\NOME_DO_SEU_ARQUIVO**.xlsx"

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    Write-Host "Abrindo arquivo..."
    $workbook = $excel.Workbooks.Open($arquivo)

    Write-Host "Atualizando conexões..."
    $workbook.RefreshAll()
    $excel.CalculateUntilAsyncQueriesDone()

    Start-Sleep -Seconds 10

    Write-Host "Registrando data e hora..."

    $workbook.Worksheets.Item(1).Range("A1").Value = (
        Get-Date
    ).ToString("dd/MM/yyyy HH:mm:ss")

    Write-Host "Salvando..."
    $workbook.Save()
    $workbook.Close($false)
}
catch {
    Write-Host "Erro: $($_.Exception.Message)"
    exit 1
}
finally {
    if ($null -ne $workbook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
            $workbook
        ) | Out-Null
    }

    if ($null -ne $excel) {
        $excel.Quit()

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
            $excel
        ) | Out-Null
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "Finalizado!"