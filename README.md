# Atualização Automática de Arquivos Excel com PowerShell e Agendador de Tarefas do Windows

Automatize a atualização de arquivos Excel sem precisar abrir a planilha manualmente. Utilizando **PowerShell**, **Microsoft Excel** e o **Agendador de Tarefas do Windows**, é possível atualizar conexões de dados, salvar alterações e registrar a data e a hora da execução.

## Objetivo

Este projeto apresenta uma solução para automatizar a atualização de planilhas Excel no Windows por meio de um script PowerShell (`.ps1`) e do Agendador de Tarefas.

### Funcionalidades

- Abrir o Excel em segundo plano.
- Abrir automaticamente a planilha desejada.
- Atualizar conexões de dados.
- Registrar a data e a hora da execução.
- Salvar as alterações realizadas.
- Fechar a planilha e encerrar o Excel.
- Agendar a execução automática em horários definidos.

## Pré-requisitos

- Windows 10 ou Windows 11.
- Microsoft Excel instalado.
- PowerShell disponível.
- Permissão para executar scripts PowerShell.
- Planilha Excel com as conexões de dados configuradas.

## 1. Organizar os arquivos

1. Abra o Explorador de Arquivos.
2. Acesse a unidade `D:` (aqui informe a unidade que está usando, no exemplo foi informado a unidade D).
3. Crie uma pasta chamada `**NOME_DA_SUA_PASTA**` (exemplo: `AutomatizarExcel`).
4. Copie a planilha `**NOME_DO_SEU_ARQUIVO**.xlsx` para essa pasta (exemplo: `CDI mensal.xlsx`).

**Importante:**

- Utilize os caminhos reais dos arquivos.
- Evite pastas com acentos ou caracteres especiais, se possível.
- Atualize os caminhos no script caso utilize outro diretório.

## 2. Autorizar a execução de scripts PowerShell

1. Abra o menu Iniciar.
2. Pesquise por PowerShell.
3. Execute o seguinte comando:

```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

4. Se solicitado, digite `S` e pressione Enter.

## 3. Criar o script PowerShell

Abra o Bloco de Notas e salve o código abaixo como `**NOME_SEU_SCRIPT**.ps1` (exemplo: `AtualizaExcel_CDI.ps1`).

```powershell
Write-Host "Iniciando Excel..."

$excel = $null
$workbook = $null

try {
    $arquivo = "D:\"**NOME_DA_SUA_PASTA\NOME_DO_SEU_ARQUIVO**.xlsx"

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    Write-Host "Abrindo arquivo..."
    $workbook = $excel.Workbooks.Open($arquivo)

    Write-Host "Atualizando conexões..."
    $workbook.RefreshAll()
    $excel.CalculateUntilAsyncQueriesDone()

    Start-Sleep -Seconds 10

    Write-Host "Registrando data e hora..."

    $workbook.Worksheets.Item(1).Range("A1").Value = (
        Get-Date
    ).ToString("dd/MM/yyyy HH:mm:ss")

    Write-Host "Salvando..."
    $workbook.Save()
    $workbook.Close($false)
}
catch {
    Write-Host "Erro: $($_.Exception.Message)"
    exit 1
}
finally {
    if ($null -ne $workbook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
            $workbook
        ) | Out-Null
    }

    if ($null -ne $excel) {
        $excel.Quit()

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
            $excel
        ) | Out-Null
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "Finalizado!"
```

## 4. Salvar e testar o script

1. Salve o arquivo em `D:\AutomatizarExcel`.
2. Pressione `Win + R`.
3. Digite `powershell`.
4. Execute o comando:

```powershell
& "D:\AutomatizarExcel\AtualizaExcel_CDI.ps1"
```

5. Abra a planilha e confira se a célula `A1` foi atualizada.
6. Verifique se as conexões de dados foram atualizadas corretamente.

## 5. Configurar o Agendador de Tarefas

1. Pressione `Win + R`.
2. Digite `taskschd.msc`.
3. Clique em **Criar Tarefa Básica**.
4. Defina o nome `Atualizar CDI Mensal`.
5. Escolha a frequência e o horário da execução.
6. Selecione **Iniciar um programa**.

**Programa/script:**

```text
powershell.exe
```

**Adicionar argumentos:**

```text
-NoProfile -ExecutionPolicy Bypass -File "D:\AutomatizarExcel\AtualizaExcel_CDI.ps1"
```

**Iniciar em:**

```text
D:\AutomatizarExcel
```

7. Conclua a configuração.
8. Execute a tarefa manualmente para testar.

## 6. Validar a automação

- [ ] O script executa sem erros.
- [ ] A planilha é aberta corretamente.
- [ ] As conexões são atualizadas.
- [ ] A data e a hora são registradas.
- [ ] A planilha é salva.
- [ ] O Excel é encerrado.
- [ ] A tarefa funciona pelo Agendador.
- [ ] A execução ocorre no horário programado.

## Solução de problemas

### Arquivo não encontrado

- Confirme o caminho da planilha.
- Verifique se o arquivo foi renomeado ou movido.

### Dados não atualizados

- Verifique as conexões no Excel.
- Confirme o acesso às fontes de dados.
- Verifique se são necessárias credenciais ou VPN.
- Avalie se o tempo de espera é suficiente.

### Tarefa não executada

- Confira o histórico do Agendador.
- Verifique o horário configurado.
- Confirme as permissões do usuário.
- Verifique se o computador estava ligado.

## Estrutura do repositório

```text
atualizacao-automatica-excel/
├── README.md
├── AtualizaExcel_CDI.ps1
└── .gitignore
```

## Considerações finais

- Mantenha uma cópia de segurança da planilha.
- Evite utilizar a célula `A1` se ela já contiver informações importantes.
- Não publique planilhas com dados confidenciais.
- Teste a atualização das conexões antes de confiar na execução automática.

## Tecnologias utilizadas

- PowerShell
- Microsoft Excel
- Agendador de Tarefas do Windows

## 👨‍💻 Autor
## Mairon Oliveira

- LinkedIn: [Mairon Oliveira](https://www.linkedin.com/in/maironhenriquedeoliveira/)
- Email: maironhenrique@hotmail.com

### 🚀 Projetos e Portfólio
Desenvolvo projetos práticos envolvendo análise de dados, dashboards, indicadores de desempenho, ETL e visualização analítica, conectando tecnologia e problemas reais de negócio.

- 📊 [Portfólio com análises no Medium](https://medium.com/@maironhenrique)
- 📈 [Portfólio com Dashboards](https://app.xperiun.com/in/maironhenrique)
