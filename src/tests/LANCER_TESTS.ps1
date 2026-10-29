param([string]$Rapport = 'RESULTATS.txt', [switch]$Capture)
$ErrorActionPreference = 'Stop'
$compilerPath = Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
$appPath = Split-Path $PSScriptRoot -Parent
$testExe = Join-Path $PSScriptRoot 'Tests.exe'
& $compilerPath /nologo /target:exe /platform:x64 "/win32manifest:$(Join-Path $appPath 'Korecto.manifest')" /main:CorrecteurMail.TestProgram "/out:$testExe" "/win32icon:$(Join-Path $appPath 'assets\Korecto.ico')" "/resource:$(Join-Path $appPath 'assets\Korecto.ico'),CorrecteurMail.App.ico" "/resource:$(Join-Path $appPath 'assets\Korecto-notification.ico'),CorrecteurMail.Notification.ico" "/resource:$(Join-Path $appPath 'assets\Korecto-128.png'),CorrecteurMail.Logo.png" "/resource:$(Join-Path $appPath 'CONSIGNES.txt'),CorrecteurMail.CONSIGNES.txt" "/resource:$(Join-Path $appPath 'response.schema.json'),CorrecteurMail.response.schema.json" "/resource:$(Join-Path $appPath 'NOUVEAUTES.txt'),CorrecteurMail.NOUVEAUTES.txt" /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:System.Web.Extensions.dll /reference:Microsoft.CSharp.dll (Join-Path $appPath 'CorrecteurMail.cs') (Join-Path $appPath 'AiConfiguration.cs') (Join-Path $appPath 'AppSettings.cs') (Join-Path $appPath 'Installation.cs') (Join-Path $appPath 'Updater.cs') (Join-Path $appPath 'SelectionAccess.cs') (Join-Path $appPath 'MailFeatures.cs') (Join-Path $PSScriptRoot 'Tests.cs')
if ($LASTEXITCODE -ne 0) { throw 'Compilation des tests échouée.' }
$reportPath = Join-Path $PSScriptRoot $Rapport
$testArguments = '"' + $reportPath + '"'
if ($Capture) { $testArguments += ' --capture' }
$testProcess = Start-Process -FilePath $testExe -ArgumentList $testArguments -WorkingDirectory $appPath -WindowStyle Hidden -PassThru
if (-not $testProcess.WaitForExit(90000)) { $testProcess.Kill(); throw 'Les tests dépassent 90 secondes.' }
Get-Content -LiteralPath $reportPath
if ($testProcess.ExitCode -ne 0) { throw 'Des tests ont échoué : consulter le rapport.' }

