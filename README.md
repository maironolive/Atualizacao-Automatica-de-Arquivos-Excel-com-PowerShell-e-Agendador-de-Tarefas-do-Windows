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
2. Acesse a unidade `D:`.
3. Crie uma pasta chamada `AutomatizarExcel`.
4. Copie a planilha `CDI mensal.xlsx` para essa pasta.

Estrutura inicial:

```text
D:\AutomatizarExcel\
└── CDI mensal.xlsx


Importante:

Utilize os caminhos reais dos seus arquivos.
Evite, se possível, pastas com acentos ou caracteres especiais.
Caso utilize outro diretório, atualize os caminhos no script e no Agendador de Tarefas.
2. Autorizar a execução de scripts PowerShell

O Windows pode restringir a execução de arquivos .ps1 por meio da política de execução do PowerShell.

Passo a passo
Abra o menu Iniciar.
Pesquise por PowerShell.
Abra o PowerShell.
Execute o comando:
powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser

Se aparecer uma solicitação de confirmação, digite S e pressione Enter.
Feche o PowerShell.

O parâmetro -Scope CurrentUser aplica a configuração ao usuário atual.

Nota de segurança: a política RemoteSigned é menos permissiva que Bypass. Em computadores corporativos, políticas definidas pela organização podem prevalecer sobre as configurações do usuário.

3. Criar o script PowerShell

Abra o Bloco de Notas e copie o código abaixo.

O script inicia o Excel em segundo plano, abre a planilha, solicita a atualização das conexões, registra a data e a hora na célula A1, salva o arquivo e encerra o Excel.

Código PowerShell
powershell
Write-Host "Iniciando Excel..."

$excel = $null
$workbook = $null

try {
    $arquivo = "D:\AutomatizarExcel\CDI mensal.xlsx"

    # Inicia o Excel em segundo plano
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    Write-Host "Abrindo arquivo..."
    $workbook = $excel.Workbooks.Open($arquivo)

    Write-Host "Atualizando conexões..."
    $workbook.RefreshAll()
    $excel.CalculateUntilAsyncQueriesDone()

    # Aguarda alguns segundos para finalizar o processamento
    Start-Sleep -Seconds 10

    Write-Host "Registrando data e hora..."

    $workbook.Worksheets.Item(1).Range("A1").Value = (
        Get-Date
    ).ToString("dd/MM/yyyy HH:mm:ss")

    Write-Host "Salvando..."
    $workbook.Save()

    Write-Host "Fechando arquivo..."
    $workbook.Close($false)
}
catch {
    Write-Host "Erro: $($_.Exception.Message)"
    exit 1
}
finally {
    if ($null -ne $workbook) {
        try {
            $workbook.Close($false)
        }
        catch {
            Write-Host "Erro ao fechar a planilha."
        }

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
            $workbook
        ) | Out-Null

        $workbook = $null
    }

    if ($null -ne $excel) {
        try {
            $excel.Quit()
        }
        catch {
            Write-Host "Erro ao encerrar o Excel."
        }

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject(
            $excel
        ) | Out-Null

        $excel = $null
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "Finalizado!"

Entendendo o código
Excel.Application: inicia o Excel por automação COM.
$excel.Visible = $false: mantém a janela do Excel oculta.
$workbook.RefreshAll(): solicita a atualização das conexões de dados.
$excel.CalculateUntilAsyncQueriesDone(): aguarda a conclusão de determinadas consultas assíncronas.
Start-Sleep -Seconds 10: adiciona uma espera de 10 segundos.
Range("A1").Value: registra a data e a hora na primeira planilha.
$workbook.Save(): salva as alterações.
$workbook.Close($false): fecha a pasta de trabalho.
$excel.Quit(): solicita o encerramento do Excel.
ReleaseComObject(): libera as referências aos objetos COM utilizados.

Importante: o tempo de espera pode variar conforme a origem dos dados e o volume de informações. Uma espera fixa de 10 segundos não garante que todas as conexões tenham terminado de atualizar.

4. Salvar o arquivo PS1

Depois de copiar o código:

No Bloco de Notas, clique em Arquivo > Salvar como.
Em Nome do arquivo, digite AtualizaExcel_CDI.ps1.
Em Tipo, selecione Todos os arquivos.
Em Codificação, selecione UTF-8, se disponível.
Salve o arquivo em D:\AutomatizarExcel.
Confirme que o arquivo não foi salvo como .ps1.txt.

Estrutura final:

text
D:\AutomatizarExcel\
├── CDI mensal.xlsx
└── AtualizaExcel_CDI.ps1

5. Testar o script manualmente

Antes de configurar o agendamento, execute o script manualmente.

Passo a passo
Pressione Win + R.
Digite powershell e pressione Enter.
Execute o comando:
powershell
& "D:\AutomatizarExcel\AtualizaExcel_CDI.ps1"

Aguarde o término da execução.
Verifique as mensagens exibidas no PowerShell.
Validar o resultado
Abra a planilha CDI mensal.xlsx.
Verifique se a célula A1 da primeira planilha contém a data e a hora da execução.
Confira se os dados das conexões foram efetivamente atualizados.
Se a célula A1 já contiver informações importantes, altere o endereço no script.
6. Verificar o encerramento do Excel

É importante confirmar que o script não deixa processos do Excel abertos indevidamente.

Passo a passo
Pressione Ctrl + Shift + Esc para abrir o Gerenciador de Tarefas.
Localize o processo Microsoft Excel ou EXCEL.EXE.
Execute novamente o script.
Aguarde a mensagem Finalizado!.
Verifique se o processo criado pela automação foi encerrado.

Atenção: não encerre indiscriminadamente todos os processos EXCEL.EXE, pois podem existir outras planilhas abertas pelo usuário.

7. Configurar o Agendador de Tarefas

O Agendador de Tarefas permite executar o script automaticamente em horários definidos.

7.1. Abrir o Agendador
Pressione Win + R.
Digite taskschd.msc.
Pressione Enter.
No painel direito, clique em Criar Tarefa Básica.
7.2. Definir a tarefa

Preencha os campos:

Nome: Atualizar CDI Mensal
Descrição: Atualiza a planilha CDI mensal automaticamente

Clique em Avançar.

7.3. Configurar o disparador

Escolha a frequência da execução:

Diariamente.
Semanalmente.
Mensalmente.
Uma única vez.

Para uma atualização diária:

Selecione Diariamente.
Defina a data e a hora da primeira execução.
Configure a recorrência desejada.
Clique em Avançar.
7.4. Configurar a ação
Selecione Iniciar um programa.
Clique em Avançar.
Preencha os campos abaixo.

Programa/script

text
powershell.exe


Adicionar argumentos

text
-NoProfile -ExecutionPolicy Bypass -File "D:\AutomatizarExcel\AtualizaExcel_CDI.ps1"


Iniciar em

text
D:\AutomatizarExcel

Clique em Avançar.
Revise as informações.
Clique em Concluir.

O Windows utilizará o PowerShell para executar o script no horário configurado.

Nota de segurança: -ExecutionPolicy Bypass aplica-se à execução iniciada por esse comando. Não utilize essa opção para contornar políticas de segurança impostas pela organização.

8. Ajustar as opções da tarefa
Abra o Agendador de Tarefas.
Localize Atualizar CDI Mensal.
Clique duas vezes na tarefa.
Acesse a aba Geral.
Para os testes iniciais, selecione Executar somente quando o usuário estiver conectado.
Clique em OK para salvar.
Executar a tarefa após um horário perdido

Para permitir que o Windows tente executar uma tarefa cujo horário tenha sido perdido:

Abra as propriedades da tarefa.
Acesse a aba Configurações.
Marque Executar a tarefa o mais cedo possível após uma execução agendada ser perdida, se disponível.
Salve as alterações.

Importante:

Essa opção não garante a execução enquanto o computador estiver desligado.
A automação COM do Excel pode apresentar limitações quando executada sem uma sessão interativa de usuário.
9. Testar a tarefa agendada

Antes de depender do agendamento, teste a tarefa diretamente.

Passo a passo
Abra o Agendador de Tarefas.
Selecione Atualizar CDI Mensal.
No painel direito, clique em Executar.
Aguarde a execução.
Abra a planilha CDI mensal.xlsx.
Confira a data e a hora registradas na célula A1.
Verifique se as conexões foram atualizadas.
Confira se o processo criado pela automação foi encerrado.

Se o script funcionar manualmente, mas não pelo Agendador, verifique:

O usuário configurado para executar a tarefa.
Os caminhos dos arquivos.
As permissões de acesso.
A disponibilidade das fontes de dados.
O histórico do Agendador de Tarefas.
10. Validar a automação

Utilize o checklist para confirmar o funcionamento.

O script .ps1 executa sem erros.
A planilha é aberta corretamente.
As conexões de dados são atualizadas.
A data e a hora são registradas na célula configurada.
A planilha é salva.
O Excel é encerrado ao final da execução.
A tarefa funciona quando executada manualmente pelo Agendador.
A tarefa é executada no horário programado.
Os dados permanecem corretos após a atualização.

Após validar todos os itens, a automação estará pronta para uso nas condições testadas.

Solução de problemas
Erro: arquivo não encontrado

Confirme se o caminho informado no script corresponde ao local real da planilha.

powershell
$arquivo = "D:\AutomatizarExcel\CDI mensal.xlsx"


Verifique também se o arquivo foi renomeado ou movido.

Erro: execução de scripts desabilitada

Consulte a política de execução:

powershell
Get-ExecutionPolicy -List


Se necessário, revise a configuração do usuário e as políticas administradas pela organização.

A planilha abre, mas os dados não são atualizados

Verifique:

Se as conexões estão configuradas corretamente no Excel.
Se a fonte de dados está acessível.
Se são necessárias credenciais ou autenticação.
Se a consulta depende de uma conexão de rede ou VPN.
Se o tempo de espera é suficiente.
Se as conexões estão configuradas para atualização em segundo plano.

Uma mensagem indicando que o script terminou não comprova, isoladamente, que todas as fontes de dados foram atualizadas com sucesso.

A célula A1 não é atualizada

Confirme se a primeira planilha é a desejada e se a célula não está protegida.

O comando utilizado é:

powershell
$workbook.Worksheets.Item(1).Range("A1").Value = (
    Get-Date
).ToString("dd/MM/yyyy HH:mm:ss")


Se necessário, substitua A1 pelo endereço correto.

O Excel continua aberto após a execução
Verifique se houve uma exceção durante o fechamento.
Confira se há outras instâncias do Excel abertas.
Consulte as mensagens do PowerShell para identificar possíveis falhas.
A tarefa não executa no horário programado

Revise:

O histórico da tarefa no Agendador de Tarefas.
A data e a hora do disparador.
O usuário configurado para executar a tarefa.
As permissões de acesso à planilha e às fontes de dados.
A disponibilidade do computador no horário programado.
As configurações de energia e suspensão do Windows.
Considerações importantes
Backup: mantenha uma cópia de segurança da planilha antes de automatizar alterações.
Célula de controle: o exemplo utiliza A1 para registrar a última execução. Escolha uma célula livre se a planilha já contiver dados nesse endereço.
Conexões externas: algumas fontes exigem autenticação, acesso à rede ou configurações adicionais.
Tempo de atualização: o tempo necessário varia conforme a quantidade de dados e o tipo de conexão.
Excel em uso: evite executar o script simultaneamente com outra rotina que modifique a mesma planilha.
Segurança: utilize apenas scripts confiáveis e configurações compatíveis com as políticas do ambiente.
Execução não interativa: a automação COM do Excel pode apresentar limitações quando executada sem uma sessão de usuário interativa.
Possíveis melhorias

Este projeto pode ser expandido com:

Registro de logs em arquivos .txt.
Envio de notificações em caso de falha.
Validação da conclusão da atualização antes de salvar.
Configuração de múltiplas planilhas no mesmo script.
Tratamento de erros mais detalhado.
Monitoramento das execuções pelo Agendador de Tarefas.
Estrutura do repositório
text
atualizacao-automatica-excel/
├── README.md
├── AtualizaExcel_CDI.ps1
└── .gitignore

Utilize o arquivo .gitignore para excluir arquivos temporários e logs, conforme necessário.
Evite publicar planilhas com dados confidenciais, credenciais ou informações pessoais.
Licença

Defina uma licença para o projeto conforme a finalidade e as condições de uso desejadas.

Tecnologias utilizadas: PowerShell, Microsoft Excel e Agendador de Tarefas do Windows.
