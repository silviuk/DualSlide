param([string]$Source,[string]$Main,[string]$Output)
$ErrorActionPreference='Stop'
$compiler=Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$wpf=Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\WPF'
& $compiler /nologo /win32manifest:"$PSScriptRoot\app.manifest" /target:exe /platform:x64 "/main:$Main" "/out:$Output" /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:Microsoft.CSharp.dll /r:System.Xaml.dll "/r:$wpf\PresentationFramework.dll" "/r:$wpf\PresentationCore.dll" "/r:$wpf\WindowsBase.dll" DualSlide.cs ModernWindows.cs PointerRenderer.cs DeckHistory.cs SessionSafety.cs $Source
if($LASTEXITCODE -ne 0){throw 'Test build failed'}
