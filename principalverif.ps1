# Coloquei no caminho local para fins de testes, então nesta versão do script há estes dois métodos de carregar o módulo da função que fiz
# Caminho local do módulo
$moduloNome = "scriptverif.psm1"
$moduloPath = Join-Path -Path $PSScriptRoot -ChildPath $moduloNome

# URL do módulo no GitHub
$moduloUrl = "https://raw.githubusercontent.com/Anazatar/verifmod/463b5473a4ab3e1d0309063e886fa05a4b0fef4f/scriptverif.psm1"

# Verifica se o módulo existe localmente; caso contrário, baixa da internet
if (-not (Test-Path $moduloPath)) {
    Write-Host "Módulo não encontrado localmente. Baixando de $moduloUrl ..." -ForegroundColor Yellow
    try {
        Invoke-WebRequest -Uri $moduloUrl -OutFile $moduloPath -UseBasicParsing
        Write-Host "Módulo baixado com sucesso." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao baixar o módulo: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
}

try {
    Import-Module $moduloPath -Force
    Write-Host "Módulo importado com sucesso." -ForegroundColor Green
} catch {
    Write-Host "Falha ao importar o módulo." -ForegroundColor Red
    return
}

# Instala dependência SPO, o ficheiro de funções
try {
    Instalar-ModuloSPO
} catch {
    return
}


$tenant = Read-Host "Digite o nome do seu tenant (sem '.onmicrosoft.com')"

$relatorioLimitacoesAplicaveis = @()
$relatorioLimitacoesNaoAplicaveis = @()

Verificar-LimitacoesTenant -tenant $tenant `
    -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis) `
    -relatorioNaoAplicaveis ([ref]$relatorioLimitacoesNaoAplicaveis)

Verificar-OneDriveSync -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis) `
    -relatorioNaoAplicaveis ([ref]$relatorioLimitacoesNaoAplicaveis)

Verificar-OneNote -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis) `
    -relatorioNaoAplicaveis ([ref]$relatorioLimitacoesNaoAplicaveis)

Exibir-Relatorios -aplicaveis $relatorioLimitacoesAplicaveis `
                  -naoAplicaveis $relatorioLimitacoesNaoAplicaveis
