param([switch]$Uninstall)
$ErrorActionPreference = 'Stop'
if (Get-Process WINWORD -ErrorAction SilentlyContinue) { throw 'Please save your documents and close Word before installing or uninstalling.' }
$guid = '{757C1B15-D248-4DC4-827B-BFD9D12B49A1}'
$progId = 'WordVocabulary.Connect'
# Register both registry views so 32-bit and 64-bit Office can load the AnyCPU assembly.
$views = @([Microsoft.Win32.RegistryView]::Registry32)
if ([Environment]::Is64BitOperatingSystem) { $views += [Microsoft.Win32.RegistryView]::Registry64 }
if (!$Uninstall) {
    & (Join-Path $PSScriptRoot 'build.ps1')
    $dll = (Resolve-Path (Join-Path $PSScriptRoot 'bin\WordVocabulary.dll')).Path
    $assembly = [Reflection.AssemblyName]::GetAssemblyName($dll).FullName
}
foreach ($view in $views) {
    $hive = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
    try {
        $paths = @("Software\Classes\CLSID\$guid", "Software\Classes\$progId", "Software\Microsoft\Office\Word\Addins\$progId")
        if ($Uninstall) { foreach ($path in $paths) { $hive.DeleteSubKeyTree($path, $false) }; continue }
        $key = $hive.CreateSubKey($paths[0]); $key.SetValue('', 'WordVocabulary.Connect'); $key.Dispose()
        $key = $hive.CreateSubKey($paths[0] + '\InprocServer32')
        $key.SetValue('', 'mscoree.dll'); $key.SetValue('ThreadingModel','Both'); $key.SetValue('Class','WordVocabulary.Connect'); $key.SetValue('Assembly',$assembly); $key.SetValue('RuntimeVersion','v4.0.30319'); $key.SetValue('CodeBase',([Uri]$dll).AbsoluteUri); $key.Dispose()
        $key = $hive.CreateSubKey($paths[0] + '\ProgId'); $key.SetValue('', $progId); $key.Dispose()
        $key = $hive.CreateSubKey($paths[1] + '\CLSID'); $key.SetValue('', $guid); $key.Dispose()
        $key = $hive.CreateSubKey($paths[2]); $key.SetValue('FriendlyName','Word Vocabulary'); $key.SetValue('Description','Click words to look up and collect vocabulary'); $key.SetValue('LoadBehavior',3,[Microsoft.Win32.RegistryValueKind]::DWord); $key.Dispose()
    } finally { $hive.Dispose() }
}
if ($Uninstall) { Write-Host 'Uninstalled. Your vocabulary and credentials were kept.' } else { Write-Host 'Installed for this Windows user. Keep this folder in place. Open Word to use the new tab.' }
