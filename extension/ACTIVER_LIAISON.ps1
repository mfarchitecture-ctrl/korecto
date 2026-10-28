param([switch]$Desactiver, [switch]$VerifierSeulement)
$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path $PSScriptRoot -Parent
$exe = Join-Path $projectRoot 'dist\Korecto.exe'
$manifestPath = Join-Path $projectRoot 'dist\com.korecto.selection.json'
$extension = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'manifest.json') -Raw | ConvertFrom-Json
$sha = [Security.Cryptography.SHA256]::Create()
try { $bytes = $sha.ComputeHash([Convert]::FromBase64String($extension.key)) } finally { $sha.Dispose() }
$id = -join (0..15 | ForEach-Object { [char](97 + ($bytes[$_] -shr 4)); [char](97 + ($bytes[$_] -band 15)) })
if (-not (Test-Path -LiteralPath $exe)) { throw 'Compilez d abord la version interne.' }
$keys = @('Software\Google\Chrome\NativeMessagingHosts\com.korecto.selection', 'Software\BraveSoftware\Brave-Browser\NativeMessagingHosts\com.korecto.selection')
if ($VerifierSeulement) { Write-Output "Liaison vérifiée pour $id ; aucune écriture registre."; exit }
if ($Desactiver) {
 foreach ($key in $keys) {
  $existing = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($key)
  $owned = $false
  if ($null -ne $existing) { try { $owned = $existing.GetValue('') -eq $manifestPath } finally { $existing.Dispose() } }
  if ($owned) { [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKey($key, $false) }
 }
 Write-Output 'Liaison de test désactivée.'; exit
}
$hostJson = @{name='com.korecto.selection';description='Sélection vers Korecto';path=$exe;type='stdio';allowed_origins=@("chrome-extension://$id/")} | ConvertTo-Json -Depth 3
[IO.File]::WriteAllText($manifestPath, $hostJson, [Text.UTF8Encoding]::new($false))
foreach ($key in $keys) {
 $existing = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($key)
 if ($null -ne $existing) { try { $old = $existing.GetValue(''); if ($old -and $old -ne $manifestPath) { throw 'Une autre liaison Korecto existe déjà : elle a été conservée.' } } finally { $existing.Dispose() } }
}
foreach ($key in $keys) { $entry = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($key); try { $entry.SetValue('', $manifestPath) } finally { $entry.Dispose() } }
Write-Output 'Liaison Chrome et Brave activée pour cette version interne.'
