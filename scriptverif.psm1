
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


#Aplicativos Personalizados de Terceiros
function Verificar-Redirect308 {
    param (
        [string]$url,
        [ref]$relatorioAplicaveis,
        [ref]$relatorioNaoAplicaveis
    )

    try {
        Invoke-WebRequest -Uri $url -MaximumRedirection 0 -ErrorAction Stop | Out-Null
        # Se não lançar exceção, então não houve redirect
        $relatorioNaoAplicaveis.Value += [PSCustomObject]@{
            Aplicativo     = "HTTP Redirect"
            Limitacao      = "Sem redirecionamento HTTP 308 detectado"
            AcaoNecessaria = "Nenhuma"
            Impacto        = "N/A"
        }
    } catch {
        if ($_.Exception.Response.StatusCode.value__ -eq 308) {
            $relatorioAplicaveis.Value += [PSCustomObject]@{
                Aplicativo     = "HTTP Redirect"
                Limitacao      = "Resposta HTTP 308 detectada"
                AcaoNecessaria = "Garantir suporte a 308 nos aplicativos"
                Impacto        = "Médio"
            }
        } else {
            $relatorioAplicaveis.Value += [PSCustomObject]@{
                Aplicativo     = "HTTP Redirect"
                Limitacao      = "Erro HTTP inesperado: $($_.Exception.Response.StatusCode)"
                AcaoNecessaria = "Edite aplicativos personalizados para garantir que eles manipulem corretamente as respostas HTTP 308."
                Impacto        = "Médio"
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

function Verificar-Delve {
    param([ref]$relatorioAplicaveis)

    $relatorioAplicaveis.Value += [PSCustomObject]@{
        Aplicativo     = "Delve"
        Limitacao      = "Pode levar 24h para exibir perfis de pessoas após a renomeação"
        AcaoNecessaria = "Nenhuma"
        Impacto        = "Médio"
    }
}


function Verificar-eDiscovery {
    param([ref]$relatorioAplicaveis)

    $relatorioAplicaveis.Value += [PSCustomObject]@{
        Aplicativo     = "Descoberta Eletrônica (eDiscovery)"
        Limitacao      = "Retenções não podem ser removidas até atualizar URLs"
        AcaoNecessaria = "Atualizar as URLs de retenção no portal do Microsoft Purview"
        Impacto        = "Médio"
    }
}

function Buscar-FormulariosInfoPathGraph {
    param (
        [string]$tenant,
        [ref]$relatorioAplicaveis,
        [ref]$relatorioNaoAplicaveis
    )

    try {
        $site = Get-MgSite -SiteId "root" -ErrorAction Stop

        if (-not $site.Id) {
            Write-Host "Erro: site raiz não retornou ID válido." -ForegroundColor Red
            return
        }

        Write-Host "Site raiz carregado: $($site.WebUrl)" -ForegroundColor Green

        $drives = Get-MgSiteDrive -SiteId $site.Id -ErrorAction Stop

        if (-not $drives) {
            Write-Host "Nenhum drive encontrado." -ForegroundColor Yellow
            return
        }

        foreach ($drive in $drives) {
            if (-not $drive.Id) {
                Write-Host "Drive sem ID detectado. Ignorando..." -ForegroundColor Yellow
                continue
            }

            Write-Host "Verificando arquivos .xsn no drive: $($drive.Name)" -ForegroundColor Cyan

            try {
                $itens = Get-MgDriveRootChild -DriveId $drive.Id -ErrorAction Stop
                $formularios = $itens | Where-Object { $_.Name -like "*.xsn" }

                if (-not $formularios) {
                    $relatorioNaoAplicaveis.Value += [PSCustomObject]@{
                        Aplicativo     = "Formulário InfoPath ($($drive.Name))"
                        Limitacao      = "Nenhum arquivo .xsn encontrado"
                        AcaoNecessaria = "Nenhuma"
                        Impacto        = "N/A"
                    }
                }

                foreach ($form in $formularios) {
                    $relatorioAplicaveis.Value += [PSCustomObject]@{
                        Aplicativo     = "Formulário InfoPath ($($drive.Name))"
                        Limitacao      = "Arquivo .xsn localizado: $($form.Name)"
                        AcaoNecessaria = "Reconectar o formulário InfoPath ao novo domínio após a renomeação"
                        Impacto        = "Médio"
                    }
                    Write-Host "InfoPath encontrado: $($form.Name)" -ForegroundColor Yellow
                }

            } catch {
                Write-Host "Erro ao acessar arquivos no drive $($drive.Name): $($_.Exception.Message)" -ForegroundColor Red
            }
        }

    } catch {
        Write-Host "Erro geral ao acessar o Graph: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Verificar-Loop {
    param([ref]$relatorioAplicaveis)

    $relatorioAplicaveis.Value += [PSCustomObject]@{
        Aplicativo     = "Microsoft Loop"
        Limitacao      = "Áreas de trabalho existentes não podem ser partilhadas ou modificadas"
        AcaoNecessaria = "Não existe nenhuma ação disponível"
        Impacto        = "Médio"
    }
}

function Verificar-SitesArquivados {
    param (
        [ref]$relatorioAplicaveis,
        [ref]$relatorioNaoAplicaveis
    )

    try {
        $sites = Get-SPOSite -Limit All | Where-Object { $_.LockState -eq "ReadOnly" }

        if ($sites.Count -gt 0) {
            $relatorioAplicaveis.Value += [PSCustomObject]@{
                Aplicativo     = "Arquivo do Microsoft 365"
                Limitacao      = "Detectado site(s) com LockState 'ReadOnly' (possivelmente arquivados)"
                AcaoNecessaria = "Reativar os sites antes da renomeação e evitar arquivar durante o processo"
                Impacto        = "Médio"
            }

            foreach ($site in $sites) {
                Write-Host "Site arquivado detectado: $($site.Url)" -ForegroundColor Yellow
            }
        } else {
            $relatorioNaoAplicaveis.Value += [PSCustomObject]@{
                Aplicativo     = "Arquivo do Microsoft 365"
                Limitacao      = "Nenhum site arquivado encontrado"
                AcaoNecessaria = "Nenhuma ação"
                Impacto        = "N/A"
            }
        }

    } catch {
        Write-Host "Erro ao verificar sites arquivados: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Verificar-MicrosoftFormsUpload {
    param (
        [ref]$relatorioAplicaveis,
        [ref]$relatorioNaoAplicaveis
    )

    $relatorioAplicaveis.Value += [PSCustomObject]@{
        Aplicativo     = "Microsoft Forms"
        Limitacao      = "Forms com campo de upload de arquivos não funcionam após renomeação"
        AcaoNecessaria = "Remover o botão de upload e adicioná-lo novamente após renomeação"
        Impacto        = "Médio"
    }

    $relatorioNaoAplicaveis.Value += [PSCustomObject]@{
        Aplicativo     = "Microsoft Forms"
        Limitacao      = "Sem campo de upload detectado"
        AcaoNecessaria = "Nenhuma"
        Impacto        = "N/A"
    }
}

function Verificar-OfficeAppsSalvamento {
    param ([ref]$relatorioAplicaveis)

    $relatorioAplicaveis.Value += [PSCustomObject]@{
        Aplicativo     = "Aplicativos do Office (Word, Excel, PowerPoint)"
        Limitacao      = "Durante a renomeação, usuários podem ter erro ao salvar arquivos hospedados"
        AcaoNecessaria = "Tentar salvar novamente ou alterar URL"
        Impacto        = "Médio"
    }
}

function Verificar-OneDriveAcessoRapido {
    param ([ref]$relatorioAplicaveis)

    $relatorioAplicaveis.Value += [PSCustomObject]@{
        Aplicativo     = "OneDrive / SharePoint - Acesso Rápido"
        Limitacao      = "Links de Acesso Rápido não funcionam após renomeação"
        AcaoNecessaria = "Usuário deve remover e recriar atalhos"
        Impacto        = "Baixo"
    }
}

function Verificar-OneDriveTeamsApp {
    param ([ref]$relatorioAplicaveis)

    $relatorioAplicaveis.Value += [PSCustomObject]@{
        Aplicativo     = "OneDrive no Teams"
        Limitacao      = "Erro 404 ao acessar OneDrive via Teams"
        AcaoNecessaria = "Enviar arquivo no chat para forçar reconfiguração"
        Impacto        = "Médio"
    }
}

function Verificar-PowerPlatformConectoresSharePoint {
    param (
        [ref]$relatorioAplicaveis,
        [ref]$relatorioNaoAplicaveis
    )

    try {
        $ambientes = Get-AdminPowerAppEnvironment
        $conectoresSP = @()

        foreach ($amb in $ambientes) {
            $conectores = Get-AdminPowerAppConnector -EnvironmentName $amb.EnvironmentName |
                          Where-Object { $_.ConnectorName -like "*sharepoint*" }

            if ($conectores) {
                $conectoresSP += $conectores
            }
        }

        if ($conectoresSP.Count -gt 0) {
            $relatorioAplicaveis.Value += [PSCustomObject]@{
                Aplicativo     = "Power Platform (Power Automate / Power BI)"
                Limitacao      = "Detectadas conexões com SharePoint Online"
                AcaoNecessaria = "Revisar fluxos e relatórios que usam URL antiga"
                Impacto        = "Alto"
            }
        } else {
            $relatorioNaoAplicaveis.Value += [PSCustomObject]@{
                Aplicativo     = "Power Platform (Power Automate / Power BI)"
                Limitacao      = "Nenhuma conexão com SharePoint detectada"
                AcaoNecessaria = "Nenhuma"
                Impacto        = "N/A"
            }
        }
    } catch {
        Write-Host "Erro ao verificar conectores Power Platform: $($_.Exception.Message)" -ForegroundColor Red

        $relatorioAplicaveis.Value += [PSCustomObject]@{
            Aplicativo     = "Power Platform"
            Limitacao      = "Falha ao obter conectores (erro ou permissões)"
            AcaoNecessaria = "Executar revisão manual nos conectores"
            Impacto        = "Médio"
        }
    }
}

function Verificar-ProjectOnlineWorkflows {
    param ([ref]$relatorioAplicaveis)

    $relatorioAplicaveis.Value += [PSCustomObject]@{
        Aplicativo     = "Project Online - Workflows"
        Limitacao      = "Workflows 'em fuga' não são concluídos e não é possível iniciar novos."
        AcaoNecessaria = "Verifique se todos os workflows foram concluídos antes da renomeação. Depois, republicar os fluxos."
        Impacto        = "Alto"
    }

    $relatorioAplicaveis.Value += [PSCustomObject]@{
        Aplicativo     = "Project Online - URLs em Workflows"
        Limitacao      = "URLs fixas nos workflows não são atualizadas com a renomeação"
        AcaoNecessaria = "Atualizar manualmente URLs em fluxos após renomeação"
        Impacto        = "Médio"
    }
}

function Verificar-ProjectOnlinePWA {
    param ([ref]$relatorioAplicaveis)

    $relatorioAplicaveis.Value += [PSCustomObject]@{
        Aplicativo     = "Project Online - PWA"
        Limitacao      = "Referências para https://project.microsoft.com deixam de funcionar"
        AcaoNecessaria = "Atualizar as URLs em 'Meu site PWA' nas definições"
        Impacto        = "Médio"
    }
}

function Verificar-ProjectOnlineExcelRelatorios {
    param ([ref]$relatorioAplicaveis)

    $relatorioAplicaveis.Value += [PSCustomObject]@{
        Aplicativo     = "Project Online - Relatórios Excel"
        Limitacao      = "Relatórios com conexões de dados do SharePoint deixam de funcionar"
        AcaoNecessaria = "Recriar conexões no Excel após a renomeação"
        Impacto        = "Alto"
    }
}

function Verificar-ProjectPro {
    param ([ref]$relatorioAplicaveis)

    $relatorioAplicaveis.Value += [PSCustomObject]@{
        Aplicativo     = "Project Pro"
        Limitacao      = "Project Pro só funciona com URL de PWA atualizada"
        AcaoNecessaria = "Atualizar URL do site PWA nas configurações da conta"
        Impacto        = "Médio"
    }
}


function Verificar-SitesHubSharePoint {
    param (
        [ref]$relatorioAplicaveis,
        [ref]$relatorioNaoAplicaveis
    )

    try {
        $hubs = Get-SPOHubSite
        if ($hubs.Count -gt 0) {
            foreach ($hub in $hubs) {
                $relatorioAplicaveis.Value += [PSCustomObject]@{
                    Aplicativo     = "Site Hub SharePoint"
                    Limitacao      = "Site Hub registrado: $($hub.SiteUrl)"
                    AcaoNecessaria = "Cancelar registro e registrar novamente se necessário após renomeação"
                    Impacto        = "Médio"
                }
            }
        } else {
            $relatorioNaoAplicaveis.Value += [PSCustomObject]@{
                Aplicativo     = "Sites Hub SharePoint"
                Limitacao      = "Nenhum site hub registrado"
                AcaoNecessaria = "Nenhuma ação"
                Impacto        = "N/A"
            }
        }
    } catch {
        Write-Host "Erro ao listar sites hub: $($_.Exception.Message)" -ForegroundColor Red
        $relatorioAplicaveis.Value += [PSCustomObject]@{
            Aplicativo     = "Sites Hub SharePoint"
            Limitacao      = "Erro na verificação"
            AcaoNecessaria = "Verificação manual recomendada"
            Impacto        = "Alto"
        }
    }
}

function Verificar-SitesBloqueados {
    param (
        [ref]$relatorioAplicaveis,
        [ref]$relatorioNaoAplicaveis
    )

    try {
        $sitesBloqueados = Get-SPOSite -Limit All | Where-Object { $_.LockState -ne "Unlock" }

        if ($sitesBloqueados.Count -gt 0) {
            foreach ($site in $sitesBloqueados) {
                $relatorioAplicaveis.Value += [PSCustomObject]@{
                    Aplicativo     = "Sites Bloqueados SharePoint/OneDrive"
                    Limitacao      = "Site bloqueado: $($site.Url) (LockState: $($site.LockState))"
                    AcaoNecessaria = "Reveja o bloqueio e remova se apropriado antes da renomeação"
                    Impacto        = "Alto"
                }
            }
        } else {
            $relatorioNaoAplicaveis.Value += [PSCustomObject]@{
                Aplicativo     = "Sites Bloqueados SharePoint/OneDrive"
                Limitacao      = "Nenhum site bloqueado detectado"
                AcaoNecessaria = "Nenhuma ação necessária"
                Impacto        = "N/A"
            }
        }
    } catch {
        Write-Host "Erro ao verificar sites bloqueados: $($_.Exception.Message)" -ForegroundColor Red
        $relatorioAplicaveis.Value += [PSCustomObject]@{
            Aplicativo     = "Sites Bloqueados SharePoint/OneDrive"
            Limitacao      = "Falha na verificação"
            AcaoNecessaria = "Verificação manual recomendada"
            Impacto        = "Médio"
        }
    }
}


function Exibir-Relatorios {
    param($aplicaveis, $naoAplicaveis)

    Write-Host "`nResumo das limitações identificadas (Aplicáveis):" -ForegroundColor Cyan
    $aplicaveis | Format-Table -AutoSize

    Write-Host "`nResumo das limitações identificadas (Não Aplicáveis):" -ForegroundColor Cyan
    $naoAplicaveis | Format-Table -AutoSize
}
