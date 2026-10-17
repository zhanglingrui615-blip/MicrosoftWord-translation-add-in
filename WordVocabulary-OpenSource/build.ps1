$ErrorActionPreference = 'Stop'
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (!(Test-Path $compiler)) { $compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe' }
$output = Join-Path $PSScriptRoot 'bin'
New-Item -ItemType Directory -Force $output | Out-Null
$sources = @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'src') -Filter '*.cs' | ForEach-Object { $_.FullName })
$refs = @('System.dll','System.Core.dll','System.Drawing.dll','System.Windows.Forms.dll','System.Net.Http.dll','System.Web.Extensions.dll','System.Security.dll','System.IO.Compression.dll','Microsoft.CSharp.dll') | ForEach-Object { '/reference:' + $_ }
& $compiler /nologo /target:library /platform:anycpu /optimize+ ('/resource:' + (Join-Path $PSScriptRoot 'data\lexicon.json') + ',WordVocabulary.Lexicon.json') ('/out:' + (Join-Path $output 'WordVocabulary.dll')) @refs @sources
if ($LASTEXITCODE -ne 0) { throw 'Compilation failed.' }
Write-Host 'Built bin\WordVocabulary.dll'
