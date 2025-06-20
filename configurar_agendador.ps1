# Criar uma nova tarefa agendada
$action = New-ScheduledTaskAction -Execute "C:\Users\eduardo.aguiar\Desktop\Automação Python itens\executar_automacao.bat"
$trigger = New-ScheduledTaskTrigger -Daily -At 8AM
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopOnIdleEnd

# Registrar a tarefa
Register-ScheduledTask -TaskName "AutomacaoExcelItensEscalada" -Action $action -Trigger $trigger -Settings $settings -Description "Executa a automação de itens de escalada diariamente às 8h" -Force 