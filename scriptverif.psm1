function Instalar-ModuloSPO {
    try {
        if (-not (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell)) {
            Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser -Force
        }
        Import-Module Microsoft.Online.SharePoint.PowerShell -Force
    }
    catch {
        Write-Host "Erro ao instalar ou importar módulo Microsoft.Online.SharePoint.PowerShell" -ForegroundColor Red
        Write-Host "Detalhes: $($_.Exception.Message)" -ForegroundColor DarkRed
        throw
    }
}

function Conectar-SharePoint {
    param([string]$adminUrl)

    try {
        Connect-SPOService -Url $adminUrl
        Write-Host "Conexão realizada com sucesso em $adminUrl" -ForegroundColor Green
        return $true
    } catch {
        Write-Host "Falha na conexão com o serviço do SharePoint Online." -ForegroundColor Red
        Write-Host "Detalhes: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Verificar-LimitacoesTenant {
    param (
        [string]$tenant,
        [ref]$relatorioAplicaveis,
        [ref]$relatorioNaoAplicaveis
    )

    $adminUrl = "https://$tenant-admin.sharepoint.com"
    if (-not (Conectar-SharePoint -adminUrl $adminUrl)) { return }

    try {
        # MultiGeo
        if ((Get-SPOTenant).MultiGeo) {
            $relatorioAplicaveis.Value += [PSCustomObject]@{
                Aplicativo     = "Tenant MultiGeo"
                Limitacao      = "Renomeação de domínio não suportada"
                AcaoNecessaria = "Não é possível prosseguir com a renomeação"
                Impacto        = "Alto"
            }
        } else {
            $relatorioNaoAplicaveis.Value += [PSCustomObject]@{
                Aplicativo     = "Tenant MultiGeo"
                Limitacao      = "Renomeação suportada"
                AcaoNecessaria = "Pode prosseguir com a renomeação"
                Impacto        = "N/A"
            }
        }

            $status = Get-SPOTenantRenameStatus
        if ($status -and $status.State -eq "InProgress") {
            $agendamento = $status.'Requested at'
            $relatorioAplicaveis.Value += [PSCustomObject]@{
                Aplicativo     = "Renomeação em Andamento"
                Limitacao      = "Renomeação já está em andamento"
                AcaoNecessaria = "Aguardar conclusão"
                Impacto        = "Alto"
            }
    } else {
            $relatorioNaoAplicaveis.Value += [PSCustomObject]@{
                Aplicativo     = "Renomeação em Andamento"
                Limitacao      = "Nenhuma renomeação ativa"
                AcaoNecessaria = "Pode prosseguir"
                Impacto        = "N/A"
        }
    }


        # Sites ativos
        $sitos = Get-SPOSite -Limit ALL | Select-Object Url, Owner
        $dominioAntigo = "$tenant.sharepoint.com"
        $comDominioAntigo = $sitos | Where-Object { $_.Url -like "*$dominioAntigo*" }
        if ($comDominioAntigo.Count -gt 0) {
            $relatorioAplicaveis.Value += [PSCustomObject]@{
                Aplicativo     = "Itens de menu do site do hub"
                Limitacao      = "URLs antigas ainda presentes"
                AcaoNecessaria = "Atualizar links manualmente"
                Impacto        = "Médio"
            }
        }

        # Sites excluídos
        $excluidos = Get-SPODeletedSite
        if ($excluidos) {
            $relatorioAplicaveis.Value += [PSCustomObject]@{
                Aplicativo     = "Sites Excluídos"
                Limitacao      = "Não restauráveis após renomeação"
                AcaoNecessaria = "Restaurar antes da mudança"
                Impacto        = "Alto"
            }
        }

        # Redirect do OneDrive
        $redirect = (Get-SPOTenant).OneDriveURLRedirect
        if ($redirect -ne "Success") {
            $relatorioAplicaveis.Value += [PSCustomObject]@{
                Aplicativo     = "OneDrive Redirect"
                Limitacao      = "Redirect não configurado corretamente"
                AcaoNecessaria = "Ativar redirect"
                Impacto        = "Médio"
            }
        } else {
            $relatorioNaoAplicaveis.Value += [PSCustomObject]@{
                Aplicativo     = "OneDrive Redirect"
                Limitacao      = "Redirect OK"
                AcaoNecessaria = "Nenhuma ação"
                Impacto        = "N/A"
            }
        }

        # Outras
        $relatorioAplicaveis.Value += @(
            @{ Aplicativo = "Office.com"; Limitacao = "Atualizações levam até 24h"; AcaoNecessaria = "Nenhuma"; Impacto = "Baixo" },
            @{ Aplicativo = "Pesquisa SharePoint"; Limitacao = "Indexação pode demorar"; AcaoNecessaria = "Aguardar"; Impacto = "Baixo" }
        ) | ForEach-Object { [PSCustomObject]$_ }

        Disconnect-SPOService
    } catch {
        Write-Host "Erro ao coletar dados do tenant: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Verificar-OneDriveSync {
    param([ref]$relatorioAplicaveis, [ref]$relatorioNaoAplicaveis)

    $path = "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe"
    if (Test-Path $path) {
        $versao = (Get-Item $path).VersionInfo.FileVersion
        if ([version]$versao -lt [version]"17.3.6943.0625") {
            $relatorioAplicaveis.Value += [PSCustomObject]@{
                Aplicativo     = "OneDrive Sync"
                Limitacao      = "Versão desatualizada"
                AcaoNecessaria = "Atualizar"
                Impacto        = "Baixo"
            }
        } else {
            $relatorioNaoAplicaveis.Value += [PSCustomObject]@{
                Aplicativo     = "OneDrive Sync"
                Limitacao      = "Versão OK"
                AcaoNecessaria = "Nenhuma"
                Impacto        = "N/A"
            }
        }
    } else {
        $relatorioAplicaveis.Value += [PSCustomObject]@{
            Aplicativo     = "OneDrive Sync"
            Limitacao      = "OneDrive.exe não encontrado"
            AcaoNecessaria = "Instalar cliente"
            Impacto        = "Baixo"
        }
    }

    $urls = @("https://oneclient.sfx.ms", "https://g.live.com")
    foreach ($url in $urls) {
        try {
            $res = Invoke-WebRequest -Uri $url -UseBasicParsing
            if ($res.StatusCode -eq 200) {
                $relatorioNaoAplicaveis.Value += [PSCustomObject]@{
                    Aplicativo     = "OneDrive - conectividade"
                    Limitacao      = "Conectividade OK"
                    AcaoNecessaria = "Nenhuma"
                    Impacto        = "N/A"
                }
            } else {
                $relatorioAplicaveis.Value += [PSCustomObject]@{
                    Aplicativo     = "OneDrive - conectividade"
                    Limitacao      = "Código HTTP $($res.StatusCode)"
                    AcaoNecessaria = "Verificar rede"
                    Impacto        = "Baixo"
                }
            }
        } catch {
            $relatorioAplicaveis.Value += [PSCustomObject]@{
                Aplicativo     = "OneDrive - conectividade"
                Limitacao      = "Erro ao acessar $url"
                AcaoNecessaria = "Permitir via firewall"
                Impacto        = "Baixo"
            }
        }
    }
}

function Verificar-OneNote {
    param([ref]$relatorioAplicaveis, [ref]$relatorioNaoAplicaveis)

    $paths = @(
        "C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE",
        "C:\Program Files (x86)\Microsoft Office\root\Office16\ONENOTE.EXE"
    )

    foreach ($p in $paths) {
        if (Test-Path $p) {
            $v = (Get-Item $p).VersionInfo.FileVersion
            if ([version]$v -lt [version]"16.0.8326.2096") {
                $relatorioAplicaveis.Value += [PSCustomObject]@{
                    Aplicativo     = "OneNote"
                    Limitacao      = "Versão desatualizada"
                    AcaoNecessaria = "Atualizar"
                    Impacto        = "Baixo"
                }
            } else {
                $relatorioNaoAplicaveis.Value += [PSCustomObject]@{
                    Aplicativo     = "OneNote"
                    Limitacao      = "Versão OK"
                    AcaoNecessaria = "Nenhuma"
                    Impacto        = "N/A"
                }
            }
            return
        }
    }

    $relatorioAplicaveis.Value += [PSCustomObject]@{
        Aplicativo     = "OneNote"
        Limitacao      = "OneNote não encontrado"
        AcaoNecessaria = "Instalar aplicativo"
        Impacto        = "Baixo"
    }
}

function Exibir-Relatorios {
    param($aplicaveis, $naoAplicaveis)

    Write-Host "`nResumo das limitações identificadas (Aplicáveis):" -ForegroundColor Cyan
    $aplicaveis | Format-Table -AutoSize

    Write-Host "`nResumo das limitações identificadas (Não Aplicáveis):" -ForegroundColor Cyan
    $naoAplicaveis | Format-Table -AutoSize
}
