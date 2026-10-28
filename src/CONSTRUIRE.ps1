param([string]$Sortie = 'Korecto.exe')
$ErrorActionPreference = 'Stop'
$compilerPath = Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
Push-Location $PSScriptRoot
try {
    $extensionResources = @('manifest.json','background.js','status.html','status.js','icon-16.png','icon-48.png','icon-128.png') | ForEach-Object { "/resource:$PSScriptRoot\..\extension\$_,CorrecteurMail.Extension.$_" }
    & $compilerPath /nologo /target:winexe /platform:x64 /win32manifest:Korecto.manifest "/out:$Sortie" /win32icon:assets/Korecto.ico /resource:assets/Korecto.ico,CorrecteurMail.App.ico /resource:assets/Korecto-notification.ico,CorrecteurMail.Notification.ico /resource:assets/Korecto-128.png,CorrecteurMail.Logo.png /resource:CONSIGNES.txt,CorrecteurMail.CONSIGNES.txt /resource:response.schema.json,CorrecteurMail.response.schema.json /resource:NOUVEAUTES.txt,CorrecteurMail.NOUVEAUTES.txt @extensionResources /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:System.Web.Extensions.dll /reference:Microsoft.CSharp.dll CorrecteurMail.cs AiConfiguration.cs AppSettings.cs Installation.cs Updater.cs SelectionAccess.cs BrowserExtension.cs
    if ($LASTEXITCODE -ne 0) { throw 'Compilation échouée.' }
    Write-Output 'Compilation terminée.'
} finally { Pop-Location }

