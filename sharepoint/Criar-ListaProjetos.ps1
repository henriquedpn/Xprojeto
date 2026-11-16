<#
.SYNOPSIS
  Cria a lista "Projetos" no SharePoint Online usando PnP PowerShell.
  Alternativa ao criar-lista-projetos.js, para quando o Console do navegador
  estiver bloqueado pela empresa. Normalmente executado pela equipe de TI.

.REQUISITOS
  - PowerShell 7.4+  e  Install-Module PnP.PowerShell -Scope CurrentUser
  - Um App Registration no Entra ID para o PnP (exigido pela Microsoft desde
    set/2024). Veja: Register-PnPEntraIDAppForInteractiveLogin
  - Permissão de Proprietário no site.

.EXEMPLO
  .\Criar-ListaProjetos.ps1 -SiteUrl https://empresa.sharepoint.com/sites/Projetos -ClientId 00000000-0000-0000-0000-000000000000
#>
param(
    [Parameter(Mandatory)] [string] $SiteUrl,
    [Parameter(Mandatory)] [string] $ClientId
)
$ErrorActionPreference = 'Stop'
$def = Get-Content -Path (Join-Path $PSScriptRoot 'definicao-lista.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$lista = $def.titulo

Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

if (Get-PnPList -Identity $lista -ErrorAction SilentlyContinue) {
    throw "Já existe uma lista '$lista' neste site. Exclua-a ou altere 'titulo' em definicao-lista.json."
}

Write-Host "► Criando a lista $lista" -ForegroundColor Cyan
New-PnPList -Title $lista -Template GenericList -EnableVersioning | Out-Null
Set-PnPList -Identity $lista -Description $def.descricao | Out-Null

Write-Host "► Criando as colunas" -ForegroundColor Cyan
$ctx = Get-PnPContext
$spList = Get-PnPList -Identity $lista
foreach ($c in $def.campos) {
    # 12 = AddToAllContentTypes + AddFieldInternalNameHint (mantém o nome interno sem acento)
    $null = $spList.Fields.AddFieldAsXml($c.xml, $false, [Microsoft.SharePoint.Client.AddFieldOptions]12)
    $ctx.ExecuteQuery()
    Write-Host "   coluna $($c.titulo)"
}

Write-Host "► Renomeando colunas e aplicando formatação" -ForegroundColor Cyan
Set-PnPField -List $lista -Identity 'Title' -Values @{ Title = $def.tituloCodigo } | Out-Null
foreach ($c in $def.campos) {
    $vals = @{ Title = $c.titulo }
    if ($c.formatador) { $vals.CustomFormatter = $c.formatador }
    Set-PnPField -List $lista -Identity $c.nome -Values $vals | Out-Null
}

Write-Host "► Configurando modos de exibição" -ForegroundColor Cyan
$padrao = Get-PnPView -List $lista | Where-Object { $_.DefaultView } | Select-Object -First 1
Set-PnPView -List $lista -Identity $padrao.Id -Fields $def.camposView `
    -Values @{ ViewQuery = $def.consultaPadrao; CustomFormatter = $def.formatadorLinha; RowLimit = 100 } | Out-Null

foreach ($v in $def.views) {
    $nova = Add-PnPView -List $lista -Title $v.titulo -Fields $def.camposView -Query $v.query -RowLimit 100 -Paged
    $vals = @{ CustomFormatter = $def.formatadorLinha }
    if ($v.totais) { $vals.Aggregations = $def.totais; $vals.AggregationsStatus = 'On' }
    Set-PnPView -List $lista -Identity $nova.Id -Values $vals | Out-Null
    Write-Host "   $($v.titulo)"
}

Write-Host "► Organizando o formulário em seções" -ForegroundColor Cyan
try {
    $ct = Get-PnPContentType -List $lista | Where-Object { $_.StringId -like '0x0100*' } | Select-Object -First 1
    $ct.ClientFormCustomFormatter = $def.formulario
    $ct.Update($false)
    Invoke-PnPQuery
} catch {
    Write-Warning "Layout do formulário não aplicado (opcional): $($_.Exception.Message)"
}

Write-Host "CONCLUÍDO! $SiteUrl/Lists/$lista" -ForegroundColor Green
