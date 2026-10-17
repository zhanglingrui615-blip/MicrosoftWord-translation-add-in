$ErrorActionPreference = 'Stop'
& (Join-Path $PSScriptRoot 'build.ps1')
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (!(Test-Path $compiler)) { $compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe' }
$bin = Join-Path $PSScriptRoot 'bin'
& $compiler /nologo /target:exe ('/out:' + (Join-Path $bin 'PublicRoutingTests.exe')) ('/reference:' + (Join-Path $bin 'WordVocabulary.dll')) /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:System.Web.Extensions.dll (Join-Path $PSScriptRoot 'tests\PublicRoutingTests.cs')
if ($LASTEXITCODE -ne 0) { throw 'Test compilation failed.' }
& (Join-Path $bin 'PublicRoutingTests.exe') $bin
if ($LASTEXITCODE -ne 0) { throw 'Tests failed.' }
& $compiler /nologo /target:exe ('/out:' + (Join-Path $bin 'BookAndExportTests.exe')) ('/reference:' + (Join-Path $bin 'WordVocabulary.dll')) /reference:Microsoft.CSharp.dll /reference:System.Drawing.dll /reference:System.Windows.Forms.dll /reference:System.IO.Compression.dll (Join-Path $PSScriptRoot 'tests\BookAndExportTests.cs')
if ($LASTEXITCODE -ne 0) { throw 'Book/export test compilation failed.' }
& (Join-Path $bin 'BookAndExportTests.exe') $bin
if ($LASTEXITCODE -ne 0) { throw 'Book/export tests failed.' }

