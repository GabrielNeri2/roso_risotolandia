<#
 =====================================================================
  Robô de reservas - Risotolândia (formulário Solvis)
  ---------------------------------------------------------------
  Lê a planilha robo_risotolandia.xlsx (na mesma pasta) e faz uma
  reserva no formulário https://web.solvis.net.br/s/reserva_clear_neo
  para cada nome da planilha, usando as opções de cada linha
  (empresa, prato, refeição, horário).

  * Usa somente recursos nativos do Windows (PowerShell + HTTPS).
    Não baixa nem executa nada externo — amigável para antivírus/EDR.
  * Usa o "modo de compatibilidade" do próprio formulário Solvis,
    que funciona sem navegador e sem JavaScript.
  * Na sexta-feira, quando o formulário pergunta "Sua reserva é para:",
    seleciona "Segunda-feira" (coluna "Se for na sexta" da planilha).

  Uso:
    robo_risotolandia.cmd    -> envia as reservas de verdade
    testar_sem_enviar.cmd    -> percorre o formulário mas NÃO finaliza
 =====================================================================
#>
param(
    # Modo teste: preenche tudo mas NÃO clica em "Finalizar"
    [switch]$SomenteTeste,

    # Caminho da planilha (padrão: robo_risotolandia.xlsx ao lado do script)
    [string]$Planilha = '',

    # Endereço do formulário
    [string]$Url = 'https://web.solvis.net.br/s/reserva_clear_neo'
)

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

# Habilita TLS 1.2 (necessário no Windows PowerShell 5.1 antigo)
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
} catch { }

$PastaScript = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
if (-not $Planilha) { $Planilha = Join-Path $PastaScript 'robo_risotolandia.xlsx' }
$LogArquivo = Join-Path $PastaScript ('log_robo_{0}.txt' -f (Get-Date -Format 'yyyy-MM-dd'))
$UrlBase    = ([uri]$Url).GetLeftPart([UriPartial]::Authority)

# ------------------------------------------------------------------ utilidades

function Escrever-Log {
    param([string]$Mensagem, [ConsoleColor]$Cor = [ConsoleColor]::Gray)
    $linha = '{0}  {1}' -f (Get-Date -Format 'HH:mm:ss'), $Mensagem
    Write-Host $linha -ForegroundColor $Cor
    Add-Content -Path $LogArquivo -Value $linha -Encoding UTF8
}

# Remove acentos e espaços e deixa minúsculo, para comparar textos
function Normalizar {
    param([string]$Texto)
    if (-not $Texto) { return '' }
    $decomposto = $Texto.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($c in $decomposto.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($c) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($c)
        }
    }
    return ($sb.ToString() -replace '\s+', '').ToLowerInvariant()
}

# Similaridade entre dois textos (0 a 1) - tolera pequenas diferenças de
# digitação entre a planilha e o site (ex.: "Selezione" x "Selezioni")
function Obter-Similaridade {
    param([string]$A, [string]$B)
    if ($A -eq $B) { return 1.0 }
    if (-not $A -or -not $B) { return 0.0 }
    $la = $A.Length; $lb = $B.Length
    $anterior = 0..$lb
    for ($i = 1; $i -le $la; $i++) {
        $atual = @($i) + (@(0) * $lb)
        for ($j = 1; $j -le $lb; $j++) {
            $custo = if ($A[$i - 1] -eq $B[$j - 1]) { 0 } else { 1 }
            $atual[$j] = [Math]::Min([Math]::Min($atual[$j - 1] + 1, $anterior[$j] + 1), $anterior[$j - 1] + $custo)
        }
        $anterior = $atual
    }
    return 1.0 - ($anterior[$lb] / [Math]::Max($la, $lb))
}

# ------------------------------------------------------------------ planilha

# Lê a planilha e devolve uma lista de pessoas com suas opções.
# 1º tenta abrir o .xlsx pelo Excel instalado; se não der, procura um .csv
# com o mesmo nome (salvo pelo Excel como "CSV (separado por vírgulas)").
function Ler-Planilha {
    param([string]$Caminho)

    $pessoas = New-Object System.Collections.Generic.List[object]

    function Mapear-Coluna {
        param($Registro, [string]$NomeColuna, [string]$ValorCelula)
        switch -Wildcard (Normalizar $NomeColuna) {
            '*empresa*'  { $Registro.Empresa      = $ValorCelula }
            '*unidade*'  { $Registro.Empresa      = $ValorCelula }
            '*prato*'    { $Registro.Prato        = $ValorCelula }
            '*refeicao*' { $Registro.Refeicao     = $ValorCelula }
            '*horario*'  { $Registro.Horario      = $ValorCelula }
            '*sexta*'    { $Registro.SeForNaSexta = $ValorCelula }
            '*nome*'     { $Registro.Nome         = $ValorCelula }
        }
    }

    function Novo-Registro {
        [pscustomobject]@{
            Empresa = ''; Prato = ''; Refeicao = ''; Horario = ''
            SeForNaSexta = ''; Nome = ''
        }
    }

    # --- tentativa 1: Excel via COM (requer Excel instalado) ---
    if (Test-Path $Caminho) {
        $excel = $null
        try { $excel = New-Object -ComObject Excel.Application } catch { $excel = $null }
        if ($excel) {
            $pasta = $null
            try {
                $excel.Visible = $false
                $excel.DisplayAlerts = $false
                $pasta = $excel.Workbooks.Open((Resolve-Path $Caminho).Path, 0, $true)
                $aba = $pasta.Worksheets.Item(1)
                $totalLinhas  = $aba.UsedRange.Rows.Count
                $totalColunas = $aba.UsedRange.Columns.Count
                $cabecalhos = @{}
                for ($c = 1; $c -le $totalColunas; $c++) {
                    $cabecalhos[$c] = [string]$aba.Cells.Item(1, $c).Text
                }
                for ($r = 2; $r -le $totalLinhas; $r++) {
                    $registro = Novo-Registro
                    for ($c = 1; $c -le $totalColunas; $c++) {
                        $valor = ([string]$aba.Cells.Item($r, $c).Text).Trim()
                        Mapear-Coluna $registro $cabecalhos[$c] $valor
                    }
                    if ($registro.Nome) { $pessoas.Add($registro) }
                }
            } finally {
                if ($pasta) { $pasta.Close($false) }
                $excel.Quit()
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
                [GC]::Collect()
            }
            return $pessoas
        }
    }

    # --- tentativa 2: arquivo .csv com o mesmo nome ---
    $caminhoCsv = [IO.Path]::ChangeExtension($Caminho, '.csv')
    if (Test-Path $caminhoCsv) {
        $primeiraLinha = Get-Content -Path $caminhoCsv -TotalCount 1
        $separador = if (($primeiraLinha -split ';').Count -gt ($primeiraLinha -split ',').Count) { ';' } else { ',' }
        foreach ($linha in (Import-Csv -Path $caminhoCsv -Delimiter $separador)) {
            $registro = Novo-Registro
            foreach ($coluna in $linha.PSObject.Properties) {
                Mapear-Coluna $registro $coluna.Name ([string]$coluna.Value).Trim()
            }
            if ($registro.Nome) { $pessoas.Add($registro) }
        }
        return $pessoas
    }

    throw "Não foi possível ler a planilha '$Caminho'. Verifique se o arquivo existe e se o Excel está instalado (ou salve uma cópia como .csv na mesma pasta)."
}

# ------------------------------------------------------------------ formulário

# Extrai do HTML da página os dados necessários para responder a etapa atual
function Extrair-Pagina {
    param([string]$Html)
    $pagina = @{ Acao = ''; Token = ''; GrupoId = ''; Pergunta = ''; Opcoes = @(); CampoTexto = ''; Botao = 'Avançar' }

    $m = [regex]::Match($Html, '<form[^>]*action="([^"]+)"'); if ($m.Success) { $pagina.Acao = $m.Groups[1].Value }
    $m = [regex]::Match($Html, 'name="authenticity_token"\s+value="([^"]+)"'); if ($m.Success) { $pagina.Token = $m.Groups[1].Value }
    $m = [regex]::Match($Html, 'name="first_group_question_id"[^>]*value="([^"]+)"'); if ($m.Success) { $pagina.GrupoId = $m.Groups[1].Value }

    $m = [regex]::Match($Html, "<h1 class=.question-text.>\s*([\s\S]*?)\s*</h1>")
    if ($m.Success) {
        $pagina.Pergunta = ([regex]::Replace($m.Groups[1].Value, '<[^>]+>', ' ') -replace '\s+', ' ').Trim()
    }

    $opcoes = New-Object System.Collections.Generic.List[object]
    $regexOpcao = '<input type="radio" name="(answers_params\[\d+\]\[choice_id\])"[^>]*value="(\d+)"[^>]*/>\s*<div>\s*<label[^>]*>\s*([\s\S]*?)\s*</label>'
    foreach ($r in [regex]::Matches($Html, $regexOpcao)) {
        $opcoes.Add([pscustomobject]@{
            Campo = $r.Groups[1].Value
            Valor = $r.Groups[2].Value
            Texto = ([regex]::Replace($r.Groups[3].Value, '<[^>]+>', ' ') -replace '\s+', ' ').Trim()
        })
    }
    $pagina.Opcoes = $opcoes

    $m = [regex]::Match($Html, '<input type="text" name="(answers_params\[\d+\]\[text\])"'); if ($m.Success) { $pagina.CampoTexto = $m.Groups[1].Value }
    $m = [regex]::Match($Html, 'name="submit_confirm_button"\s+value="([^"]+)"'); if ($m.Success) { $pagina.Botao = $m.Groups[1].Value }

    return $pagina
}

# Monta o corpo do POST no formato de formulário web (codificação UTF-8)
function Montar-Corpo {
    param([hashtable]$Campos)
    $partes = foreach ($item in $Campos.GetEnumerator()) {
        '{0}={1}' -f [uri]::EscapeDataString([string]$item.Key), [uri]::EscapeDataString([string]$item.Value)
    }
    return ($partes -join '&')
}

# Dado o texto da pergunta, devolve o(s) valor(es) da planilha a procurar
function Valores-Para-Pergunta {
    param([string]$Pergunta, $Pessoa)
    $q = Normalizar $Pergunta
    if ($q -like '*unidade*' -or $q -like '*empresa*') { return @($Pessoa.Empresa) }
    if ($q -like '*prato*')                            { return @($Pessoa.Prato) }
    if ($q -like '*refeicao*')                         { return @($Pessoa.Refeicao) }
    if ($q -like '*horario*')                          { return @($Pessoa.Horario) }
    # Pergunta que só aparece na sexta-feira: "Sua reserva é para:"
    if ($q -like '*reservaepara*' -or $q -like '*sabado*') {
        $dia = if ($Pessoa.SeForNaSexta) { $Pessoa.SeForNaSexta } else { 'Segunda-feira' }
        return @($dia)
    }
    # Pergunta não reconhecida: tenta todos os valores da linha
    return @($Pessoa.Empresa, $Pessoa.Prato, $Pessoa.Refeicao, $Pessoa.Horario, $Pessoa.SeForNaSexta) | Where-Object { $_ }
}

# Escolhe, entre as opções da tela, a que melhor combina com os valores desejados
function Escolher-Opcao {
    param($Opcoes, [string[]]$Desejados)
    $melhor = $null; $melhorNota = 0.0
    foreach ($opcao in $Opcoes) {
        $textoOpcao = Normalizar $opcao.Texto
        foreach ($desejado in $Desejados) {
            $textoDesejado = Normalizar $desejado
            if (-not $textoDesejado) { continue }
            $nota = 0.0
            if ($textoOpcao -eq $textoDesejado) { $nota = 1.0 }
            elseif ($textoOpcao.Contains($textoDesejado) -or $textoDesejado.Contains($textoOpcao)) { $nota = 0.95 }
            else { $nota = Obter-Similaridade $textoOpcao $textoDesejado }
            if ($nota -gt $melhorNota) { $melhorNota = $nota; $melhor = $opcao }
        }
    }
    return [pscustomobject]@{ Opcao = $melhor; Nota = $melhorNota }
}

# Percorre o formulário inteiro para uma pessoa. Devolve $true se deu certo.
function Fazer-Reserva {
    param($Pessoa, [switch]$Teste)

    $sessao = $null
    $resposta = Invoke-WebRequest -Uri ($Url + '?compatibility_mode=true') -SessionVariable sessao -UseBasicParsing -TimeoutSec 90 `
        -UserAgent 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) RoboRisotolandia'

    $perguntaAnterior = ''
    for ($etapa = 1; $etapa -le 12; $etapa++) {
        $pagina = Extrair-Pagina $resposta.Content
        if (-not $pagina.Acao -or -not $pagina.Token) {
            throw 'Página inesperada: o formulário não foi encontrado.'
        }
        $destino = if ($pagina.Acao -match '^https?://') { $pagina.Acao } else { $UrlBase + $pagina.Acao }

        # ---- última etapa: campo de nome ----
        if ($pagina.CampoTexto) {
            if ($Teste) {
                Escrever-Log ("   [TESTE] Chegou até '{0}' — nome '{1}' NÃO foi enviado." -f $pagina.Pergunta, $Pessoa.Nome) Yellow
                return $true
            }
            $campos = @{
                authenticity_token      = $pagina.Token
                browser_fingerprint     = ''
                first_group_question_id = $pagina.GrupoId
                submit_confirm_button   = $pagina.Botao
            }
            $campos[$pagina.CampoTexto] = $Pessoa.Nome
            $resposta = Invoke-WebRequest -Uri $destino -Method Post -Body (Montar-Corpo $campos) `
                -ContentType 'application/x-www-form-urlencoded; charset=UTF-8' `
                -WebSession $sessao -UseBasicParsing -TimeoutSec 90
            if ($resposta.Content -match 'question-text') {
                $sobra = (Extrair-Pagina $resposta.Content).Pergunta
                Escrever-Log ("   ATENÇÃO: após finalizar ainda apareceu a pergunta '{0}'. Confira manualmente." -f $sobra) Red
                return $false
            }
            Escrever-Log ("   OK: reserva finalizada para {0}." -f $Pessoa.Nome) Green
            return $true
        }

        # ---- etapa de múltipla escolha ----
        if ($pagina.Opcoes.Count -eq 0) {
            throw ("Etapa sem opções e sem campo de texto (pergunta: '{0}')." -f $pagina.Pergunta)
        }
        if ($pagina.Pergunta -eq $perguntaAnterior) {
            throw ("O site repetiu a pergunta '{0}' — resposta não aceita." -f $pagina.Pergunta)
        }
        $perguntaAnterior = $pagina.Pergunta

        $desejados = @(Valores-Para-Pergunta $pagina.Pergunta $Pessoa)
        $escolha = Escolher-Opcao $pagina.Opcoes $desejados
        if (-not $escolha.Opcao -or $escolha.Nota -lt 0.75) {
            $disponiveis = ($pagina.Opcoes | ForEach-Object { $_.Texto }) -join ' | '
            throw ("Nenhuma opção combinou com '{0}' na pergunta '{1}'. Opções do site: {2}" -f ($desejados -join ', '), $pagina.Pergunta, $disponiveis)
        }
        Escrever-Log ("   {0} -> {1}" -f $pagina.Pergunta, $escolha.Opcao.Texto)

        $campos = @{
            authenticity_token      = $pagina.Token
            browser_fingerprint     = ''
            first_group_question_id = $pagina.GrupoId
            submit_confirm_button   = $pagina.Botao
        }
        $campos[$escolha.Opcao.Campo] = $escolha.Opcao.Valor
        $resposta = Invoke-WebRequest -Uri $destino -Method Post -Body (Montar-Corpo $campos) `
            -ContentType 'application/x-www-form-urlencoded; charset=UTF-8' `
            -WebSession $sessao -UseBasicParsing -TimeoutSec 90
    }
    throw 'O formulário tem mais etapas do que o esperado.'
}

# ------------------------------------------------------------------ execução

Escrever-Log '================ Robô de reservas Risotolândia ================' Cyan
Escrever-Log ("Formulário: {0}" -f $Url)
Escrever-Log ("Planilha:   {0}" -f $Planilha)
if ($SomenteTeste) {
    Escrever-Log 'MODO TESTE: o robô vai percorrer o formulário mas NÃO vai finalizar nenhuma reserva.' Yellow
}

$pessoas = @(Ler-Planilha $Planilha)
if ($pessoas.Count -eq 0) {
    Escrever-Log 'Nenhum nome encontrado na planilha. Nada a fazer.' Yellow
    exit 0
}
Escrever-Log ("{0} pessoa(s) na planilha." -f $pessoas.Count)

$sucessos = 0; $falhas = 0
foreach ($pessoa in $pessoas) {
    Escrever-Log ("-> {0} ({1} / {2} / {3} / {4})" -f $pessoa.Nome, $pessoa.Empresa, $pessoa.Prato, $pessoa.Refeicao, $pessoa.Horario) White
    try {
        if (Fazer-Reserva -Pessoa $pessoa -Teste:$SomenteTeste) { $sucessos++ } else { $falhas++ }
    } catch {
        Escrever-Log ("   ERRO para {0}: {1}" -f $pessoa.Nome, $_.Exception.Message) Red
        $falhas++
    }
    Start-Sleep -Seconds 2   # pausa educada entre uma pessoa e outra
}

$corFinal = if ($falhas -gt 0) { [ConsoleColor]::Red } else { [ConsoleColor]::Green }
Escrever-Log ("Concluído: {0} com sucesso, {1} com falha." -f $sucessos, $falhas) $corFinal
Escrever-Log ("Registro salvo em: {0}" -f $LogArquivo)
if ($falhas -gt 0) { exit 1 } else { exit 0 }
