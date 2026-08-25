<#
.SYNOPSIS
    Script de Auditoria Forense para busca de termos em documentos do Office, txt, etc.
.PARAMETER Termos
    Lista de nomes ou strings para busca (Ex: "NOME1", "NOME2").

.PARAMETER Caminho
    Diretorio inicial da varredura. O padrao e a pasta atual (.).

.EXAMPLE
    .\Search-OpenXmlNoOffice.ps1 -Termos ".crypt", "ransom", "bitcoin" -Caminho "%userprofile%\Downloads"
#>

param([string[]]$Termos, [string]$Caminho = ".")

$Pattern = $Termos -join "|"
$Files = Get-ChildItem -Path $Caminho -Include *.docx, *.xlsx, *.pptx -Recurse

foreach ($file in $Files) {
    $tempDir = Join-Path $env:TEMP "ForensicScan_$(Get-Random)"
    try {
        # Renomeia temporariamente para .zip e extrai
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($file.FullName, $tempDir)

        # Busca em todos os arquivos XML extraídos
        $match = Get-ChildItem -Path $tempDir -Recurse -Filter *.xml | Select-String -Pattern $Pattern

        if ($match) {
            Write-Output "[!] Termo encontrado via extração XML em: $($file.FullName)"
        }
    } catch {
        Write-Warning "Falha ao descompactar ou analisar XML do arquivo $($file.FullName): $($_.Exception.Message)"
    }
    finally { if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force } }
}