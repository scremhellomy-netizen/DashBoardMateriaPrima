# Dashboard Matéria Prima — atualização local + publicação no GitHub Pages.
# Execute pelo arquivo INICIAR_ATUALIZACAO.cmd, na mesma pasta.
# Requisitos: Windows, Excel desktop, Git, internet e permissão para publicar no repositório.
$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$ExcelPath = Join-Path $Root 'Query - Dados - Lari.xlsx'
$IqyPath = Join-Path $Root 'query - Lari - Novas Colunas.iqy'
$JsonPath = Join-Path $Root 'data\materia_prima.json'
$HtmlPath = Join-Path $Root 'Dash_Board_Materia_Prima.html'
$SheetName = 'query - Lari'
$Remote = 'https://github.com/scremhellomy-netizen/DashBoardMateriaPrima.git'
$Required = @('ID','Criado','ID ESCALA','Nome Solicitante','Processo Solicitacao','Quantidade','Ação da Solicitação','Status Final','DATA DE ATEND ESTOQUE','DATA DE FINALIZAÇÃO','qtd atendida')
$DateColumns = @('Criado','DATA DE ATEND ESTOQUE','DATA DE FINALIZAÇÃO')
$excel = $null; $workbook = $null; $sheet = $null; $range = $null; $snapshot = $null
function Invoke-Git {
    param([string[]]$Arguments)
    & git @Arguments
    if ($LASTEXITCODE -ne 0) { throw "Git falhou: git $($Arguments -join ' ') (código $LASTEXITCODE)" }
}
try {
    Set-Location -LiteralPath $Root
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) { throw 'Git não encontrado. Instale Git for Windows.' }
    if (-not (Test-Path -LiteralPath $IqyPath)) { throw "Fonte oficial IQY não encontrada: $IqyPath" }
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
    $remoteOutput = @(& git remote get-url origin 2>&1)
    if ($LASTEXITCODE -ne 0 -or $remoteOutput.Count -eq 0) { throw 'Não foi possível obter o remoto origin. Verifique a configuração do Git.' }
    $remoteActual = ([string]($remoteOutput | Select-Object -First 1)).Trim()
    if ([string]::IsNullOrWhiteSpace($remoteActual)) { throw 'O remoto origin retornou um endereço vazio.' }
    if ($remoteActual -ne $Remote) { throw "A pasta pertence a outro repositório: $remoteActual. Extraia o pacote em uma pasta nova, sem .git, e execute novamente." }
    $branchOutput = @(& git branch --show-current 2>&1)
    if ($LASTEXITCODE -ne 0) { throw 'Não foi possível identificar a branch Git.' }
    $branch = "main"
    if ($branch -ne 'main') { Invoke-Git -Arguments @('checkout','-B','main') }

    Write-Host 'Consultando a NOVA fonte oficial SharePoint (.iqy)...' -ForegroundColor Cyan
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    # O arquivo IQY é a fonte primária. A planilha XLSX será apenas uma cópia local dos dados retornados.
    # Na primeira execução, o Excel pode solicitar login/permissão para acessar o SharePoint.
    $workbook = $excel.Workbooks.Open($IqyPath,0,$true)
    if ($null -eq $workbook) { throw 'Excel não abriu a conexão IQY.' }
    foreach ($ws in $workbook.Worksheets) { foreach ($qt in $ws.QueryTables) { try { $qt.BackgroundQuery = $false } catch {} } }
    try { $workbook.RefreshAll() } catch { throw "Falha ao atualizar consulta IQY: $($_.Exception.Message)" }
    $deadline = (Get-Date).AddMinutes(10)
    do {
        try { $excel.CalculateUntilAsyncQueriesDone() } catch { }
        Start-Sleep -Seconds 2
        $busy = $false
        foreach ($connection in $workbook.Connections) {
            try { if ($connection.OLEDBConnection.Refreshing) { $busy = $true } } catch { }
            try { if ($connection.ODBCConnection.Refreshing) { $busy = $true } } catch { }
        }
        foreach ($ws in $workbook.Worksheets) {
            foreach ($qt in $ws.QueryTables) {
                try { if ($qt.Refreshing) { $busy = $true } } catch { }
            }
        }
        if ((Get-Date) -gt $deadline) { throw 'Tempo limite aguardando a consulta IQY. Nenhuma alteração foi publicada.' }
    } while ($busy)
    # A primeira aba da pasta aberta pelo IQY é a tabela retornada pela consulta.
    $sheet = $workbook.Worksheets.Item(1)
    $range = $sheet.UsedRange
    if ($range.Rows.Count -lt 2) { throw 'A nova consulta IQY não retornou registros. Nada foi publicado.' }
    $values = $range.Value2
    $headers = @()
    for ($c=1; $c -le $range.Columns.Count; $c++) { $headers += ([string]$values[1,$c]).Trim() }
    foreach ($name in $Required) { if ($headers -notcontains $name) { throw "Coluna obrigatória ausente na nova consulta: $name" } }
    foreach ($name in @('Prioridade','MotivoQtdDiferente','COMENTARIOS INTERNOS')) {
        if ($headers -notcontains $name) { throw "A consulta IQY não retornou a coluna esperada: $name. Nada foi publicado." }
    }
    if ($headers.Count -ne @($headers | Sort-Object -Unique).Count) {
        throw 'A consulta IQY retornou nomes de colunas duplicados. Nada foi publicado.'
    }
    Write-Host "Nova fonte validada: $($range.Rows.Count - 1) linhas e $($headers.Count) colunas." -ForegroundColor Cyan
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
    # Cria cópia de leitura local (snapshot) somente após validar todos os dados.
    # O .iqy continua sendo a única fonte de dados; o .xlsx NÃO alimenta mais o JSON.
    Write-Host 'Gravando cópia atualizada Query - Dados - Lari.xlsx...' -ForegroundColor Cyan
    $snapshot = $excel.Workbooks.Add()
    $snapshotSheet = $snapshot.Worksheets.Item(1)
    $snapshotSheet.Name = $SheetName
    $destRange = $snapshotSheet.Range($snapshotSheet.Cells.Item(1,1),$snapshotSheet.Cells.Item($range.Rows.Count,$range.Columns.Count))
    $destRange.Value2 = $values
    $snapshotTemp = Join-Path $Root 'Query - Dados - Lari.NOVO.xlsx'
    if (Test-Path -LiteralPath $snapshotTemp) { Remove-Item -LiteralPath $snapshotTemp -Force }
    $snapshot.SaveAs($snapshotTemp,51)
    $snapshot.Close($false); $snapshot = $null
    # Copia a planilha validada sobre a cópia local existente.
    # Evita File.Replace com terceiro argumento nulo no Windows PowerShell.
    Copy-Item -LiteralPath $snapshotTemp -Destination $ExcelPath -Force -ErrorAction Stop
    Remove-Item -LiteralPath $snapshotTemp -Force -ErrorAction Stop

    $payload = [ordered]@{ updatedAt = (Get-Date).ToString('o'); rows = $rows }
    $json = ConvertTo-Json -InputObject $payload -Depth 12
    # Gravação atômica para não deixar JSON truncado em caso de falha.
    $tmp = "$JsonPath.tmp"
    [System.IO.File]::WriteAllText($tmp,$json,[System.Text.UTF8Encoding]::new($false))
    # Substitui também o JSON existente sem usar File.Replace(..., $null).
    Copy-Item -LiteralPath $tmp -Destination $JsonPath -Force -ErrorAction Stop
    Remove-Item -LiteralPath $tmp -Force -ErrorAction Stop

    Write-Host 'Preparando publicação no GitHub...' -ForegroundColor Cyan
    Invoke-Git -Arguments @('add','--','Dash_Board_Materia_Prima.html','data/materia_prima.json','ATUALIZAR_MATERIA_PRIMA.ps1','INICIAR_ATUALIZACAO.cmd')
    & git diff --cached --quiet
    if ($LASTEXITCODE -eq 1) {
        Invoke-Git -Arguments @('commit','-m','Atualiza dashboard e dados de materia prima')
    } elseif ($LASTEXITCODE -ne 0) { throw 'Falha ao verificar alterações do Git.' }
    # Se houve novos commits no remoto, tenta aplicar nosso commit por cima sem sobrescrever o remoto.
    Invoke-Git -Arguments @('fetch','origin','main')
    Invoke-Git -Arguments @('pull','--rebase','origin','main')
    Invoke-Git -Arguments @('push','origin','main')
    Write-Host "SUCESSO! $unique IDs publicados. O GitHub Pages pode levar alguns minutos para atualizar." -ForegroundColor Green
    Write-Host 'Os usuários podem abrir o dashboard e clicar em Atualizar dados.' -ForegroundColor Green
} catch {
    $details = "ERRO: $($_.Exception.Message)`r`nLinha: $($_.InvocationInfo.ScriptLineNumber)`r`nComando: $($_.InvocationInfo.Line)`r`n$($_.ScriptStackTrace)"
    Write-Host $details -ForegroundColor Red
    try { $details | Out-File -FilePath (Join-Path $Root 'erro_atualizacao.log') -Encoding UTF8 } catch {}
    Write-Host 'A publicação não foi concluída. Verifique o erro antes de tentar novamente.' -ForegroundColor Yellow
    exit 1
} finally {
    if ($snapshot -ne $null) { try { $snapshot.Close($false) } catch {} }
    if ($workbook -ne $null) { try { $workbook.Close($false) } catch {} }
    if ($excel -ne $null) { try { $excel.Quit() } catch {} }
    foreach ($obj in @($range,$sheet,$workbook,$excel)) {
        if ($null -ne $obj) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch {} }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
