<#
.SYNOPSIS
    Script de Auditoria Forense para busca de termos em documentos do Office
.PARAMETER Termos
    Lista de nomes ou strings para busca (Ex: "NOME1", "NOME2").

.PARAMETER Caminho
    Diretorio inicial da varredura. O padrao e a pasta atual (.).

.EXAMPLE
    .\Search-OfficeCom.ps1 -Termos ".crypt", "ransom", "bitcoin" -Caminho "%userprofile%\Downloads" 
#>

param(
    [Parameter(Mandatory=$true)][string[]]$Termos,
    [string]$Caminho = "."
)

$Pattern = $Termos -join "|"
$Files = Get-ChildItem -Path $Caminho -Include *.doc*, *.xls*, *.ppt* -Recurse -ErrorAction SilentlyContinue

# Inicializa os motores (invisíveis)
$word = New-Object -ComObject Word.Application
$excel = New-Object -ComObject Excel.Application
$ppt = New-Object -ComObject PowerPoint.Application
$word.Visible = $excel.Visible = $false

foreach ($file in $Files) {
    $ext = $file.Extension.ToLower()
    Write-Output "Analisando: $($file.Name)"

    try {
        if ($ext -match "doc") {
            $doc = $word.Documents.Open($file.FullName, $false, $true)
            if ($doc.Content.Text -match $Pattern) { Write-Output "[!] Encontrado em Word: $($file.FullName)" }
            $doc.Close()
        }
        elseif ($ext -match "xls") {
            $wb = $excel.Workbooks.Open($file.FullName)
            foreach ($sheet in $wb.Worksheets) {
                foreach ($t in $Termos) {
                    if ($sheet.UsedRange.Find($t)) { Write-Output "[!] Encontrado em Excel: $($file.FullName)"; break }
                }
            }
            $wb.Close($false)
        }
        elseif ($ext -match "ppt") {
            $pres = $ppt.Presentations.Open($file.FullName, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
            foreach ($slide in $pres.Slides) {
                foreach ($shape in $slide.Shapes) {
                    if ($shape.HasTextFrame -and $shape.TextFrame.TextRange.Text -match $Pattern) {
                        Write-Output "[!] Encontrado em PowerPoint: $($file.FullName)"
                        break
                    }
                }
            }
            $pres.Close()
        }
    } catch { Write-Output "[Erro] Falha ao processar $($file.Name): $($_.Exception.Message)" }
}

$word.Quit(); $excel.Quit(); $ppt.Quit()