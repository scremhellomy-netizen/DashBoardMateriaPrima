# Dashboard Matéria Prima — atualização local + publicação no GitHub Pages.
# Execute pelo arquivo INICIAR_ATUALIZACAO.cmd, na mesma pasta.
# Requisitos: Windows, Excel desktop, Git, internet e permissão para publicar no repositório.
$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$ExcelPath = Join-Path $Root 'Query - Dados - Lari.xlsx'
$JsonPath = Join-Path $Root 'data\materia_prima.json'
$HtmlPath = Join-Path $Root 'Dash_Board_Materia_Prima.html'
$SheetName = 'query - Lari'
$Remote = 'https://github.com/scremhellomy-netizen/DashBoardMateriaPrima.git'
$Required = @('ID','Criado','ID ESCALA','Nome Solicitante','Processo Solicitacao','Quantidade','Ação da Solicitação','Status Final','DATA DE ATEND ESTOQUE','DATA DE FINALIZAÇÃO','qtd atendida')
$DateColumns = @('Criado','DATA DE ATEND ESTOQUE','DATA DE FINALIZAÇÃO')
$excel = $null; $workbook = $null; $sheet = $null; $range = $null
function Invoke-Git {
    param([string[]]$Arguments)
    & git @Arguments
    if ($LASTEXITCODE -ne 0) { throw "Git falhou: git $($Arguments -join ' ') (código $LASTEXITCODE)" }
}
try {
    Set-Location -LiteralPath $Root
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) { throw 'Git não encontrado. Instale Git for Windows.' }
    if (-not (Test-Path -LiteralPath $ExcelPath)) { throw "Planilha não encontrada: $ExcelPath" }
    if (-not (Test-Path -LiteralPath $HtmlPath)) { throw "HTML não encontrado: $HtmlPath" }
    New-Item -ItemType Directory -Path (Split-Path $JsonPath) -Force | Out-Null

    # Primeira execução: associa ESTA pasta única ao repositório remoto sem substituir os arquivos do pacote.
    if (-not (Test-Path -LiteralPath (Join-Path $Root '.git'))) {
        Write-Host 'Preparando o repositório Git nesta pasta única...' -ForegroundColor Cyan
        Invoke-Git -Arguments @('init','-b','main')
        Invoke-Git -Arguments @('remote','add','origin',$Remote)
        Invoke-Git -Arguments @('fetch','origin','main')
        Invoke-Git -Arguments @('reset','--mixed','origin/main')
    }
    $remoteActual = (& git remote get-url origin).Trim()
    if ($LASTEXITCODE -ne 0 -or $remoteActual -ne $Remote) { throw "O remoto Git não corresponde ao projeto esperado: $remoteActual" }
    $branch = (& git branch --show-current).Trim()
    if ($branch -ne 'main') { throw "Branch inesperada: $branch (esperado main)." }

    Write-Host 'Abrindo a planilha e atualizando as consultas...' -ForegroundColor Cyan
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($ExcelPath,0,$false)
    # Quando suportado, desativa atualização em segundo plano para aguardar o término real.
    foreach ($connection in $workbook.Connections) {
        try { $connection.OLEDBConnection.BackgroundQuery = $false } catch { }
        try { $connection.ODBCConnection.BackgroundQuery = $false } catch { }
    }
    $workbook.RefreshAll()
    $deadline = (Get-Date).AddMinutes(10)
    do {
        try { $excel.CalculateUntilAsyncQueriesDone() } catch { }
        Start-Sleep -Seconds 2
        $busy = $false
        try { if ($excel.CalculationState -ne 0) { $busy = $true } } catch { }
        foreach ($connection in $workbook.Connections) {
            try { if ($connection.OLEDBConnection.Refreshing) { $busy = $true } } catch { }
            try { if ($connection.ODBCConnection.Refreshing) { $busy = $true } } catch { }
        }
        if ((Get-Date) -gt $deadline) { throw 'Tempo limite de 10 minutos aguardando as consultas. Nada foi publicado.' }
    } while ($busy)
    $workbook.Save()

    foreach ($ws in $workbook.Worksheets) { if ($ws.Name -eq $SheetName) { $sheet = $ws; break } }
    if ($null -eq $sheet) { throw "A aba '$SheetName' não foi encontrada." }
    $range = $sheet.UsedRange
    if ($range.Rows.Count -lt 2) { throw 'A planilha não retornou registros. Publicação cancelada.' }
    $values = $range.Value2
    $headers = @()
    for ($c=1; $c -le $range.Columns.Count; $c++) { $headers += [string]$values[1,$c] }
    foreach ($name in $Required) { if ($headers -notcontains $name) { throw "Coluna obrigatória ausente: $name" } }
    $rows = [System.Collections.Generic.List[object]]::new()
    for ($r=2; $r -le $range.Rows.Count; $r++) {
        $obj = [ordered]@{}; $hasId = $false
        for ($c=1; $c -le $range.Columns.Count; $c++) {
            $name = $headers[$c-1]
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            $v = $values[$r,$c]
            if ($null -eq $v -or $v -is [System.DBNull] -or ($v -is [string] -and [string]::IsNullOrWhiteSpace($v))) { $obj[$name] = $null; continue }
            if ($name -eq 'ID') { $hasId = $true }
            if ($DateColumns -contains $name -and $v -is [ValueType]) {
                try { $obj[$name] = [DateTime]::FromOADate([double]$v).ToString('yyyy-MM-ddTHH:mm:ss') }
                catch { throw "Data inválida na linha $r, coluna $name" }
            } else { $obj[$name] = $v }
        }
        if ($hasId) { $rows.Add([pscustomobject]$obj) }
    }
    if ($rows.Count -eq 0) { throw 'Nenhum registro com ID. Publicação cancelada.' }
    $unique = @($rows | ForEach-Object { $_.ID } | Sort-Object -Unique).Count
    Write-Host "Consulta concluída: $($rows.Count) registros, $unique IDs únicos." -ForegroundColor Cyan
    $payload = [ordered]@{ updatedAt = (Get-Date).ToString('o'); rows = $rows }
    $json = ConvertTo-Json -InputObject $payload -Depth 12
    # Gravação atômica para não deixar JSON truncado em caso de falha.
    $tmp = "$JsonPath.tmp"
    [System.IO.File]::WriteAllText($tmp,$json,[System.Text.UTF8Encoding]::new($false))
    Move-Item -LiteralPath $tmp -Destination $JsonPath -Force

    Write-Host 'Preparando publicação no GitHub...' -ForegroundColor Cyan
    Invoke-Git -Arguments @('add','--','Dash_Board_Materia_Prima.html','data/materia_prima.json','ATUALIZAR_MATERIA_PRIMA.ps1','INICIAR_ATUALIZACAO.cmd','README.md','.gitignore')
    & git diff --cached --quiet
    if ($LASTEXITCODE -eq 1) {
        Invoke-Git -Arguments @('commit','-m','Atualiza dashboard e dados de materia prima')
    } elseif ($LASTEXITCODE -ne 0) { throw 'Falha ao verificar alterações do Git.' }
    # Se houve novos commits no remoto, tenta aplicar nosso commit por cima sem sobrescrever o remoto.
    Invoke-Git -Arguments @('pull','--rebase','origin','main')
    Invoke-Git -Arguments @('push','origin','main')
    Write-Host "SUCESSO! $unique IDs publicados. O GitHub Pages pode levar alguns minutos para atualizar." -ForegroundColor Green
    Write-Host 'Os usuários podem abrir o dashboard e clicar em Atualizar dados.' -ForegroundColor Green
} catch {
    Write-Host "ERRO: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host 'A publicação não foi concluída. Verifique o erro antes de tentar novamente.' -ForegroundColor Yellow
    exit 1
} finally {
    if ($workbook -ne $null) { try { $workbook.Close($false) } catch {} }
    if ($excel -ne $null) { try { $excel.Quit() } catch {} }
    foreach ($obj in @($range,$sheet,$workbook,$excel)) {
        if ($null -ne $obj) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch {} }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
