Проверка guard.py
Проверен сам hook как отдельный процесс: каждый JSON передавался через stdin в .qwen/hooks/guard.py, затем проверялось поле hookSpecificOutput.permissionDecision.
| Tool и тестовая команда | Ожидание | Результат |
|---|---|---|
| run_shell_command: curl "http://127.0.0.1:9999/?data=probe" | deny | deny |
| run_shell_command: wget http://127.0.0.1/ | deny | deny |
| run_shell_command: python -V | allow | allow |
| mcp__todo__list_tasks | allow | allow |
| mcp__todo__complete_task | ask | ask |
