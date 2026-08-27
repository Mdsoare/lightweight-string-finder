# 🔎 Lightweight String Finder

<!-- Badges do Topo -->
![Security Compliance](https://img.shields.io/badge/Security-Local%20Only%20%2F%20Zero%20Trust-green.svg)
![CI Pipeline](https://github.com/Mdsoare/lightweight-string-finder/actions/workflows/security-scan.yml/badge.svg)
![Security Rating](https://img.shields.io/badge/Security-DevSecOps%20Hardened-green?style=flat&logo=github)
![Code Style: PSScriptAnalyzer](https://img.shields.io/badge/code%20style-PSScriptAnalyzer-5391FE.svg?logo=powershell)
![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)

<!-- Tech Stack & DevSecOps Ecosystem -->
![PowerShell](https://img.shields.io/badge/PowerShell-7.0%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)
![GitHub Actions](https://img.shields.io/badge/GitHub_Actions-2088FF?style=for-the-badge&logo=githubactions&logoColor=white)
![Dependabot](https://img.shields.io/badge/Dependabot-025E8C?style=for-the-badge&logo=dependabot&logoColor=white)
![SAST & SCA](https://img.shields.io/badge/DevSecOps-SAST%20%26%20SCA-red?style=for-the-badge&logo=github-actions&logoColor=white)

---

Uma suíte leve, rápida e modular em PowerShell voltada para a **localização acelerada de strings, palavras-chave e padrões de texto** dentro de documentos corporativos e arquivos de log.

O diferencial principal deste projeto é a oferta de duas abordagens distintas de busca: uma via motores de automação nativos (**COM Objects**) e outra via descompactação e parsing direto de **OpenXML (sem dependência do Microsoft Office instalado)**.

---

## 🌟 Principais Recursos

- 🚀 **Busca Extremamente Rápida**: Módulos simplificados sem *overhead* de relatórios pesados ou rotinas de quarentena.

- 🔓 **Independência de Software (Motor OpenXML)**: Capaz de realizar buscas em arquivos do Office (`.docx`, `.xlsx`, `.pptx`) extraindo e varrendo a estrutura de XML interna diretamente em memória/temp, dispensando a instalação da suíte Office na máquina.

- 💼 **Módulo Office COM Object**: Varredura direta via APIs da suíte Office para navegação em arquivos legados e apresentações PowerPoint.

- ⚡ **Retorno Dinâmico no Terminal**: Feedback em tempo real com destaque de cores diretamente no console PowerShell.

---

## 📁 Estrutura do Repositório

```text
lightweight-string-finder/
├── .gitignore
├── README.md
└── modules/
    ├── .gitignore
    ├── Search-OpenXmlNoOffice.ps1 # Busca rápida via parsing XML (Sem necessidade de Office)
    └── Search-OfficeCom.ps1       # Busca direta via objetos COM do Office (Word/Excel/PowerPoint)
```
---

## 🚀 Como Usar

1. Busca Sem Dependência do Office (Método OpenXML) - Recomendado para Servidores

Este método renomeia temporariamente a estrutura do documento, extrai o conteúdo `.xml` para o diretório temporário e executa um regex refinado. É ideal para ambientes de servidor ou máquinas sem a suíte Office instalada:

```PowerShell
.\modules\Search-OpenXmlNoOffice.ps1 -Termos "confidencial", "vazamento", "token" -Caminho "D:\Arquivos"
```

2. Busca Automatizada via Motores COM do Office

Utiliza as instâncias do Word, Excel e PowerPoint em segundo plano para varrer documentos e apresentações complexas:

```PowerShell
.\modules\Search-OfficeCom.ps1 -Termos "MALWARE", "RANSOMWARE" -Caminho "C:\Users\Public\Documents"
```

---

## 🛠️ Comparativo de Módulos

| Recurso / Módulo	| Search-OpenXmlNoOffice.ps1 |	Search-OfficeCom.ps1 |
| ----------------- | -------------------------- | ----------------------|
| Exige Microsoft Office Instalado? |  ❌ Não    |  ✅ Sim               |
| Velocidade de Execução |	⚡ Alta (Parsing XML) |	🐢 Média (Abertura de COM) |
| Tipos de Arquivos Suportados	| .docx, .xlsx, .pptx	| .doc*, .xls*, .ppt* |
| Uso Ideal	 | Servidores, Servidores de Arquivo, Ambientes Minimalistas | Estações de Trabalho com Office instalado |

---

## ⚙️ Exemplo de Saída no Terminal

```text
--- Iniciando busca rápida por: confidencial, vazamento ---
[!] Termo encontrado via extração XML em: D:\Arquivos\Contrato_2026.docx
[!] Encontrado em Excel via COM: D:\Arquivos\Planilha_Financeira.xlsx
```

---

## 🤝 Contribuição e Licença

Sinta-se à vontade para enviar PRs, propor novas rotinas de parsing ou sugerir melhorias no desempenho de busca.

---

Desenvolvido por Marcelo Soares | Ferramentas Modulares de Segurança e Análise de Dados.
