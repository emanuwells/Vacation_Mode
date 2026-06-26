# Filtro de mensagens de commit para remover trailer Co-authored-by do Cursor.
$content = [Console]::In.ReadToEnd()
$lines = $content -split "`r?`n"
$lines | Where-Object { $_ -notmatch 'cursoragent@cursor\.com' }
