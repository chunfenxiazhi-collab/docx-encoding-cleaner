<#
  docx-clean.ps1 —— 清理 Word/WPS 文档里的"隐形垃圾"与WPS非标准标记，输出 <原名>_clean.docx
  处理内容：
    1) U+FEFF  (BOM / 零宽不换行空格)      —— 网页、AI 结果、PDF 复制常见残留
    2) U+200B~200D / U+2060 / U+180E       —— 零宽空格、零宽连接符、word joiner
    3) <w:cr/>  -> <w:br/>                 —— WPS 的软回车换成 Word 标准换行符
  只读原文件、只写新文件，不修改原件。
  用法：  .\docx-clean.ps1 "C:\path\文件.docx"      或把文件拖到 clean-docx.cmd 上
#>
[CmdletBinding()]
param(
  [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
  [string[]]$Paths
)
$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

if (-not $Paths) {
  Write-Host "用法: .\docx-clean.ps1 <文档1.docx> [文档2.docx ...]"
  Write-Host "   或把 docx 文件拖到 clean-docx.cmd 上"
  exit 1
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

foreach ($p in $Paths) {
  $src = (Resolve-Path -LiteralPath $p).Path
  if ([IO.Path]::GetExtension($src) -notmatch '^\.docx$') { Write-Host "[跳过] 不是 .docx: $src"; continue }
  $dst = Join-Path (Split-Path $src) ([IO.Path]::GetFileNameWithoutExtension($src) + '_clean.docx')

  $tmp = Join-Path $env:TEMP ('docxclean_' + [guid]::NewGuid().ToString('N'))
  New-Item -ItemType Directory -Path $tmp | Out-Null
  try {
    Copy-Item -LiteralPath $src -Destination (Join-Path $tmp 'in.zip')
    Expand-Archive -LiteralPath (Join-Path $tmp 'in.zip') -DestinationPath (Join-Path $tmp 'x') -Force

    $totalFeff = 0; $totalZw = 0; $totalCr = 0; $touched = @()
    Get-ChildItem (Join-Path $tmp 'x\word') -Filter *.xml -File -Recurse | ForEach-Object {
      $text = [IO.File]::ReadAllText($_.FullName, [Text.Encoding]::UTF8)
      $feff = ([regex]::Matches($text, '\uFEFF')).Count
      $zw = ([regex]::Matches($text, '[\u200B-\u200D\u2060\u180E]')).Count
      $cr = ([regex]::Matches($text, '<w:cr\s*/>')).Count
      if ($feff -or $zw -or $cr) {
        $text = $text -replace '\uFEFF', ''
        $text = $text -replace '[\u200B-\u200D\u2060\u180E]', ''
        $text = $text -replace '<w:cr\s*/>', '<w:br/>'
        [IO.File]::WriteAllText($_.FullName, $text, $utf8NoBom)
        $touched += ("      word\{0}: FEFF={1} 零宽={2} 软回车={3}" -f $_.Name, $feff, $zw, $cr)
      }
      $totalFeff += $feff; $totalZw += $zw; $totalCr += $cr
    }

    if (Test-Path -LiteralPath $dst) { Remove-Item -LiteralPath $dst -Force }
    [IO.Compression.ZipFile]::CreateFromDirectory((Join-Path $tmp 'x'), $dst, [IO.Compression.CompressionLevel]::Optimal, $false)

    Write-Host ("[完成] {0}" -f (Split-Path $dst -Leaf))
    Write-Host ("      U+FEFF={0}  零宽字符={1}  软回车<w:cr/>={2}" -f $totalFeff, $totalZw, $totalCr)
    if ($touched) { $touched | ForEach-Object { Write-Host $_ } }
    if (-not ($totalFeff + $totalZw + $totalCr)) { Write-Host "      未发现需要清理的标记（文件本身是干净的）" }
  }
  catch {
    Write-Host ("[失败] {0}`n      {1}" -f $src, $_.Exception.Message)
    Write-Host "      若文件正被 Word/WPS 打开，请先关闭再试。"
  }
  finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
}
