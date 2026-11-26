param(
    [string]$OutDir = (Join-Path $PSScriptRoot 'dist'),
    [string]$FileName = 'MathTypeChapSync.dotm'
)

$ErrorActionPreference = 'Stop'
$srcDir = Join-Path $PSScriptRoot 'src'
$basPath = Join-Path $srcDir 'MTChapSync.bas'
$clsPath = Join-Path $srcDir 'MTChapSyncEvents.cls'
$thisDocPath = Join-Path $srcDir 'ThisDocument.txt'
$uiPath = Join-Path $srcDir 'customUI.xml'

if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
$outFile = Join-Path $OutDir $FileName

if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
    throw 'Close all Word windows before building.'
}

# Build in a temp folder: "SaveAs to a brand-new file" hangs on this machine,
# but "Save" over an existing template file works. So seed from Normal.dotm.
$build = Join-Path $env:TEMP ('mtcs_build_' + [guid]::NewGuid().ToString('N').Substring(0,8))
New-Item -ItemType Directory -Path $build | Out-Null
$target = Join-Path $build $FileName
$normal = Join-Path $env:APPDATA 'Microsoft\Templates\Normal.dotm'
if (-not (Test-Path $normal)) { throw "Normal.dotm not found: $normal" }
Copy-Item $normal $target -Force

# VBE imports ANSI; convert UTF-8 source to the system codepage (GBK on zh-CN)
$ansi = [System.Text.Encoding]::GetEncoding([System.Globalization.CultureInfo]::CurrentCulture.TextInfo.ANSICodePage)
$basAnsi = Join-Path $build 'MTChapSync.bas'
$codeText = [System.IO.File]::ReadAllText($basPath, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText($basAnsi, $codeText, $ansi)

$docCode = [System.IO.File]::ReadAllText($thisDocPath, [System.Text.Encoding]::UTF8)

$secKey = 'HKCU:\Software\Microsoft\Office\16.0\Word\Security'
if (-not (Test-Path $secKey)) { New-Item -Path $secKey -Force | Out-Null }
$oldVbom = (Get-ItemProperty -Path $secKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
New-ItemProperty -Path $secKey -Name AccessVBOM -Value 1 -PropertyType DWord -Force | Out-Null

$word = $null
$proc = $null
try {
    # Start Word WITHOUT add-ins ( /a ): MathType/iSlide/Zotero popups hang COM automation
    $winword = 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE'
    $proc = Start-Process $winword -ArgumentList '/a' -PassThru
    Start-Sleep -Seconds 6
    $word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3

    $doc = $word.Documents.Open($target, $false, $false)

    # VBProject can be $null for a moment right after Word starts - wait for it
    $waited = 0
    while ($null -eq $doc.VBProject -and $waited -lt 30) {
        Start-Sleep -Milliseconds 500
        $waited++
    }
    if ($null -eq $doc.VBProject) { throw 'VBProject is not accessible (AccessVBOM / Trust Center).' }

    $proj = $doc.VBProject
    # 必须改掉继承自 Normal.dotm 的 'TemplateProject'，否则与 MathType 的工程同名，
    # Word 解析功能区回调时会解析到错误的工程，报 '错误的参数号或无效的属性赋值'
    $proj.Name = 'MathTypeChapSync'
    for ($i = $proj.VBComponents.Count; $i -ge 1; $i--) {
        $c = $proj.VBComponents.Item($i)
        if ($c.Name -ne 'ThisDocument') { $proj.VBComponents.Remove($c) }
    }

    $td = $proj.VBComponents.Item('ThisDocument')
    if ($td.CodeModule.CountOfLines -gt 0) { $td.CodeModule.DeleteLines(1, $td.CodeModule.CountOfLines) }
    $td.CodeModule.AddFromString($docCode)

    $null = $proj.VBComponents.Import($basAnsi)

    $names = @($proj.VBComponents | ForEach-Object { $_.Name })
    'modules: ' + ($names -join ', ')

    if ($doc.Content.Text.Length -gt 0) { $doc.Content.Delete() }

    $doc.Save()
    'saved: ' + (Get-Item $target).Length + ' bytes'
    $doc.Close(0)
} finally {
    if ($word) {
        try { $word.Quit() } catch {}
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch {}
    }
    Start-Sleep -Milliseconds 800
    if ($proc) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
    if ($null -eq $oldVbom) {
        Remove-ItemProperty -Path $secKey -Name AccessVBOM -ErrorAction SilentlyContinue
    } else {
        New-ItemProperty -Path $secKey -Name AccessVBOM -Value $oldVbom -PropertyType DWord -Force | Out-Null
    }
}

# ---- inject customUI (ribbon) ----
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression

function Set-ZipText($zip, [string]$name, [string]$text) {
    $e = $zip.GetEntry($name)
    if ($e) { $e.Delete() }
    $e = $zip.CreateEntry($name)
    $st = $e.Open()
    $sw = New-Object System.IO.StreamWriter($st, (New-Object System.Text.UTF8Encoding($false)))
    $sw.Write($text); $sw.Flush(); $sw.Dispose(); $st.Dispose()
}

function Get-ZipText($zip, [string]$name) {
    $e = $zip.GetEntry($name)
    if (-not $e) { return $null }
    $sr = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
    $t = $sr.ReadToEnd(); $sr.Dispose()
    return $t
}

$uiXml = [System.IO.File]::ReadAllText($uiPath, [System.Text.Encoding]::UTF8)
$zip = [System.IO.Compression.ZipFile]::Open($target, 'Update')
try {
    Set-ZipText $zip 'customUI/customUI.xml' $uiXml
    $rels = Get-ZipText $zip '_rels/.rels'
    if ($rels -notmatch 'customUI/customUI.xml') {
        $rel = '<Relationship Id="rIdMTChapSyncUI" Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" Target="customUI/customUI.xml"/>'
        $rels = $rels -replace '</Relationships>', ($rel + '</Relationships>')
        Set-ZipText $zip '_rels/.rels' $rels
    }
    $ct = Get-ZipText $zip '[Content_Types].xml'
    if ($ct -notmatch 'customUI/customUI.xml') {
        $ov = '<Override PartName="/customUI/customUI.xml" ContentType="application/xml"/>'
        $ct = $ct -replace '</Types>', ($ov + '</Types>')
        Set-ZipText $zip '[Content_Types].xml' $ct
    }
} finally { $zip.Dispose() }

if (Test-Path $outFile) { Remove-Item $outFile -Force }
Copy-Item $target $outFile -Force
Remove-Item $build -Recurse -Force

$zip = [System.IO.Compression.ZipFile]::OpenRead($outFile)
try {
    $entries = @($zip.Entries | ForEach-Object { $_.FullName })
    'contains customUI: ' + ($entries -contains 'customUI/customUI.xml')
    'contains vbaProject: ' + ($entries -contains 'word/vbaProject.bin')
} finally { $zip.Dispose() }

'built: ' + $outFile + ' (' + (Get-Item $outFile).Length + ' bytes)'
