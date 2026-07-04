# ============================================================
# 範本正規化腳本（開發用，一次性）
# 將原始 ODT 範本中的示範資料（紅字＝公司資料、藍字＝每案資料、
# 一處黑字公司名）替換為 ${參數} 佔位符，掛上 PARAM.* 專屬樣式，
# 並以確定性方式重新打包（mimetype 第一個 entry、stored）。
# 產出：C:\WorkSpace\tender_file\範本\*.odt ＋ tokens.txt manifest
# 原始範本唯讀不動。
#
# ※ 必須用 PowerShell 7+（pwsh）執行。Windows PowerShell 5.1 的
#   ZipArchive 即使指定 NoCompression 仍會把 mimetype 寫成 Deflate，
#   產出的範本 mimetype 不是 stored、不符 ODF 規範。
# ============================================================
#requires -Version 7.0
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null

$SrcDir = 'C:\Users\firej\OneDrive\Desktop\恩潔標案相關文件'
$OutDir = 'C:\WorkSpace\tender_file\範本'

$NS_TEXT   = 'urn:oasis:names:tc:opendocument:xmlns:text:1.0'
$NS_STYLE  = 'urn:oasis:names:tc:opendocument:xmlns:style:1.0'
$NS_FO     = 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0'
$NS_OFFICE = 'urn:oasis:names:tc:opendocument:xmlns:office:1.0'

# 範本中的示範資料字面值
$LIT_COMPANY = @('哈瑪星科技股份有限公司', '哈瑪星科技有限公司')  # 長的先比對
$LIT_BOSS    = '陳文德'
$LIT_PHONE   = '(07)536-4800'
$LIT_ADDR    = @('高雄市前鎮區民權二路', '8', '號', '18', '樓')    # 相鄰 5 span 合併
$LIT_ORG     = '花蓮縣政府'
$LIT_CASE    = @('115', '年全球資訊網子網站系統維護服務案')        # 相鄰 2 span 合併

# 殘留檢查字面值（正規化後 body 文字中必須為 0）
$RESIDUALS = @('哈瑪星', '陳文德', '536-4800', '民權二路', '花蓮縣政府',
               '全球資訊網子網站系統維護服務案', '115')

function New-NsMgr([xml]$doc) {
    $m = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $m.AddNamespace('text',   $NS_TEXT)
    $m.AddNamespace('style',  $NS_STYLE)
    $m.AddNamespace('fo',     $NS_FO)
    $m.AddNamespace('office', $NS_OFFICE)
    return ,$m   # XmlNamespaceManager 是 IEnumerable，防止 PowerShell 將其展開
}

function Get-SpanStyleName($span) { return $span.GetAttribute('style-name', $NS_TEXT) }

# 建立（或重用）PARAM 樣式：複製 base 樣式、強制紅色（範本中醒目標示參數）
function Ensure-ParamStyle([xml]$doc, $nsMgr, [hashtable]$cache, [string]$token, [string]$baseStyle) {
    $name = 'PARAM.' + $token + '.' + ($baseStyle -replace '[^0-9A-Za-z一-鿿]', '_')
    if ($cache.ContainsKey($name)) { return $name }
    $autoStyles = $doc.SelectSingleNode('//office:automatic-styles', $nsMgr)
    $base = $doc.SelectSingleNode("//office:automatic-styles/style:style[@style:name='$baseStyle']", $nsMgr)
    if ($base -ne $null) {
        $clone = $base.CloneNode($true)
    } else {
        # base 是段落預設或 styles.xml 中的樣式（例如 WW-預設段落字型）：建立僅含顏色的文字樣式
        $clone = $doc.CreateElement('style', 'style', $NS_STYLE)
        $fam = $doc.CreateAttribute('style', 'family', $NS_STYLE); $fam.Value = 'text'
        [void]$clone.Attributes.Append($fam)
        [void]$clone.AppendChild($doc.CreateElement('style', 'text-properties', $NS_STYLE))
    }
    $clone.SetAttribute('name', $NS_STYLE, $name) | Out-Null
    $tp = $clone.SelectSingleNode('style:text-properties', $nsMgr)
    if ($tp -eq $null) { $tp = $clone.AppendChild($doc.CreateElement('style', 'text-properties', $NS_STYLE)) }
    $tp.SetAttribute('color', $NS_FO, '#FF0000') | Out-Null
    [void]$autoStyles.AppendChild($clone)
    $cache[$name] = $true
    return $name
}

# 以單一 PARAM span 取代一串 span（第一個之外全部移除）
function Replace-Spans([xml]$doc, $nsMgr, [hashtable]$styleCache, [object[]]$spans, [string]$token) {
    $first = $spans[0]
    $styleName = Ensure-ParamStyle $doc $nsMgr $styleCache $token (Get-SpanStyleName $first)
    $newSpan = $doc.CreateElement('text', 'span', $NS_TEXT)
    $newSpan.SetAttribute('style-name', $NS_TEXT, $styleName) | Out-Null
    $newSpan.InnerText = '${' + $token + '}'
    [void]$first.ParentNode.ReplaceChild($newSpan, $first)
    for ($i = 1; $i -lt $spans.Count; $i++) { [void]$spans[$i].ParentNode.RemoveChild($spans[$i]) }
}

# 取得 span 的「緊鄰下一個 sibling span」（允許中間出現 bookmark 標記，不允許文字）
function Next-AdjacentSpan($span) {
    $n = $span.NextSibling
    while ($n -ne $null) {
        if ($n.NodeType -eq 'Text') { if ($n.Value.Trim() -ne '') { return $null } }
        elseif ($n.LocalName -eq 'span') { return $n }
        elseif ($n.LocalName -in @('bookmark-start', 'bookmark-end', 'bookmark')) { }  # 穿透
        else { return $null }
        $n = $n.NextSibling
    }
    return $null
}

if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Force $OutDir | Out-Null }
$manifest = New-Object System.Collections.Generic.List[string]
$allOk = $true

foreach ($srcFile in Get-ChildItem $SrcDir -Filter *.odt) {
    Write-Host "`n===== 處理 $($srcFile.Name) ====="

    # --- 讀取原始 ODT 全部 entry ---
    $entries = [ordered]@{}   # name -> byte[]
    $zip = [System.IO.Compression.ZipFile]::OpenRead($srcFile.FullName)
    try {
        foreach ($e in $zip.Entries) {
            $ms = New-Object System.IO.MemoryStream
            $s = $e.Open(); $s.CopyTo($ms); $s.Dispose()
            $entries[$e.FullName] = $ms.ToArray()
        }
    } finally { $zip.Dispose() }

    # --- 解析 content.xml ---
    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $true
    $doc.LoadXml([System.Text.Encoding]::UTF8.GetString($entries['content.xml']))
    $nsMgr = New-NsMgr $doc
    $styleCache = @{}

    # 顏色樣式表（content.xml 的 automatic styles）
    $colorOf = @{}
    foreach ($st in $doc.SelectNodes('//office:automatic-styles/style:style', $nsMgr)) {
        $tp = $st.SelectSingleNode('style:text-properties', $nsMgr)
        if ($tp -ne $null) {
            $c = $tp.GetAttribute('color', $NS_FO)
            if ($c -eq '#FF0000' -or $c -eq '#0070C0') { $colorOf[$st.GetAttribute('name', $NS_STYLE)] = $c }
        }
    }

    # --- Pass 1：span 層級字面值規則（顏色無關 → 同時涵蓋黑字公司名）---
    $spans = @($doc.SelectNodes('//office:body//text:span', $nsMgr))
    $consumed = New-Object 'System.Collections.Generic.HashSet[object]'
    foreach ($sp in $spans) {
        if ($consumed.Contains($sp) -or $sp.ParentNode -eq $null) { continue }
        $t = $sp.InnerText
        $matched = $null; $group = @($sp)

        if     ($LIT_COMPANY -contains $t) { $matched = '廠商名稱' }
        elseif ($t -eq $LIT_BOSS)          { $matched = '負責人姓名' }
        elseif ($t -eq $LIT_PHONE)         { $matched = '聯絡電話' }
        elseif ($t -eq $LIT_ORG)           { $matched = '機關名稱' }
        elseif ($t -eq $LIT_ADDR[0]) {
            # 地址 5 段合併
            $seq = @($sp); $cur = $sp; $ok = $true
            for ($i = 1; $i -lt $LIT_ADDR.Count; $i++) {
                $cur = Next-AdjacentSpan $cur
                if ($cur -eq $null -or $cur.InnerText -ne $LIT_ADDR[$i]) { $ok = $false; break }
                $seq += $cur
            }
            if ($ok) { $matched = '廠商地址'; $group = $seq }
        }
        elseif ($t -eq $LIT_CASE[0]) {
            # 案名 2 段合併（'115' 單獨出現時不在此處理，留給日期 pass）
            $nx = Next-AdjacentSpan $sp
            if ($nx -ne $null -and $nx.InnerText -eq $LIT_CASE[1]) { $matched = '案件名稱'; $group = @($sp, $nx) }
        }

        if ($matched -ne $null) {
            foreach ($g in $group) { [void]$consumed.Add($g) }
            Replace-Spans $doc $nsMgr $styleCache $group $matched
            Write-Host ("  [token] {0} <= '{1}'" -f $matched, (($group | ForEach-Object { $_.InnerText }) -join ''))
        }
    }

    # --- Pass 2：日期段（中華民國…年…月…日）藍色 run → 簽署年/月/日 ---
    foreach ($p in $doc.SelectNodes('//office:body//text:p', $nsMgr)) {
        if ($p.InnerText -notmatch '中\s*　*華\s*　*民\s*　*國') { continue }
        $blueSpans = @($p.SelectNodes('.//text:span', $nsMgr) | Where-Object {
            $colorOf[(Get-SpanStyleName $_)] -eq '#0070C0' })
        if ($blueSpans.Count -eq 0) { continue }
        # 分組：緊鄰的藍 span 併成 run
        $runs = @(); $run = @()
        foreach ($bs in $blueSpans) {
            if ($run.Count -gt 0 -and (Next-AdjacentSpan $run[-1]) -eq $bs) { $run += $bs }
            else { if ($run.Count -gt 0) { $runs += ,$run }; $run = @($bs) }
        }
        if ($run.Count -gt 0) { $runs += ,$run }
        if ($runs.Count -ne 3) { throw "日期段藍色 run 數量為 $($runs.Count)（預期 3）：$($p.InnerText)" }
        $tokens = @('簽署年', '簽署月', '簽署日')
        for ($i = 0; $i -lt 3; $i++) {
            $txt = (($runs[$i] | ForEach-Object { $_.InnerText }) -join '') -replace '[\s　]', ''
            if ($txt -notmatch '^\d+$') { throw "日期 run 非數字：'$txt'" }
            Replace-Spans $doc $nsMgr $styleCache $runs[$i] $tokens[$i]
            Write-Host ("  [token] {0} <= '{1}'" -f $tokens[$i], $txt)
        }
    }

    # --- Pass 3：殘留的紅/藍 span → 白名單（底線、空白/間距）改黑，其他報錯 ---
    foreach ($sp in @($doc.SelectNodes('//office:body//text:span', $nsMgr))) {
        $sn = Get-SpanStyleName $sp
        if (-not $colorOf.ContainsKey($sn)) { continue }
        $t = $sp.InnerText -replace '[\s　]', ''
        if ($t -eq '' -or $t -eq '_') {
            # 直接把該 automatic style 的顏色改黑（junk 樣式僅 junk span 使用）
            $st = $doc.SelectSingleNode("//office:automatic-styles/style:style[@style:name='$sn']", $nsMgr)
            $tp = $st.SelectSingleNode('style:text-properties', $nsMgr)
            $tp.SetAttribute('color', $NS_FO, '#000000') | Out-Null
            Write-Host ("  [junk->黑] style={0} text='{1}'" -f $sn, $sp.InnerText)
        } else {
            throw "未處理的彩色 span：style=$sn text='$($sp.InnerText)'"
        }
    }

    # --- Pass 3.5：移除已無任何引用的紅/藍孤兒樣式（被替換掉的 span 遺留的定義）---
    foreach ($sn in @($colorOf.Keys)) {
        $refs = $doc.SelectNodes("//*[@*[local-name()='style-name' and . = '$sn']]", $nsMgr)
        if ($refs.Count -eq 0) {
            $st = $doc.SelectSingleNode("//office:automatic-styles/style:style[@style:name='$sn']", $nsMgr)
            if ($st -ne $null) { [void]$st.ParentNode.RemoveChild($st); Write-Host "  [孤兒樣式移除] $sn" }
        }
    }

    # --- Pass 4：殘留字面值檢查（body 文字）---
    $bodyText = $doc.SelectSingleNode('//office:body', $nsMgr).InnerText
    foreach ($r in $RESIDUALS) {
        if ($bodyText.Contains($r)) { Write-Host "  [殘留!] '$r' 仍出現於 body 文字" -ForegroundColor Red; $allOk = $false }
    }
    # meta.xml / styles.xml 也掃一次（資訊性）
    foreach ($aux in @('meta.xml', 'styles.xml')) {
        $auxText = [System.Text.Encoding]::UTF8.GetString($entries[$aux])
        foreach ($r in $RESIDUALS) {
            if ($r -ne '115' -and $auxText.Contains($r)) { Write-Host "  [注意] $aux 內含 '$r'" -ForegroundColor Yellow }
        }
    }

    # --- 序列化 content.xml（UTF-8 無 BOM）---
    $ms = New-Object System.IO.MemoryStream
    $xw = [System.Xml.XmlWriter]::Create($ms, (New-Object System.Xml.XmlWriterSettings -Property @{
        Encoding = (New-Object System.Text.UTF8Encoding($false)); Indent = $false }))
    $doc.Save($xw); $xw.Dispose()
    $entries['content.xml'] = $ms.ToArray()

    # --- 確定性重打包：mimetype 第一、stored；其餘依原順序 ---
    $outPath = Join-Path $OutDir $srcFile.Name
    if (Test-Path $outPath) { Remove-Item $outPath -Force }
    $fs = [System.IO.File]::Create($outPath)
    $za = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $names = @('mimetype') + @($entries.Keys | Where-Object { $_ -ne 'mimetype' })
        foreach ($n in $names) {
            $level = if ($n -eq 'mimetype') { [System.IO.Compression.CompressionLevel]::NoCompression }
                     else { [System.IO.Compression.CompressionLevel]::Optimal }
            $entry = $za.CreateEntry($n, $level)
            $es = $entry.Open(); $es.Write($entries[$n], 0, $entries[$n].Length); $es.Dispose()
        }
    } finally { $za.Dispose(); $fs.Dispose() }

    # --- 驗證輸出 ---
    $bytes = [System.IO.File]::ReadAllBytes($outPath)
    $sig = ($bytes[0] -eq 0x50 -and $bytes[1] -eq 0x4B)
    $method = [BitConverter]::ToUInt16($bytes, 8)          # 第一個 local header 的壓縮法
    $nameLen = [BitConverter]::ToUInt16($bytes, 26)
    $firstName = [System.Text.Encoding]::ASCII.GetString($bytes, 30, $nameLen)
    if (-not ($sig -and $method -eq 0 -and $firstName -eq 'mimetype')) {
        Write-Host "  [錯誤] zip 結構：sig=$sig method=$method first=$firstName" -ForegroundColor Red; $allOk = $false
    } else { Write-Host "  [zip OK] mimetype 第一、stored" }
    # 重新開啟驗 XML 與 token 數
    $zin = [System.IO.Compression.ZipFile]::OpenRead($outPath)
    try {
        $s = $zin.GetEntry('content.xml').Open()
        $sr = New-Object System.IO.StreamReader($s, [System.Text.Encoding]::UTF8)
        $cxml = $sr.ReadToEnd(); $sr.Dispose()
        $chk = New-Object System.Xml.XmlDocument; $chk.LoadXml($cxml)   # well-formed 驗證
        $tokenCounts = @{}
        foreach ($m in [regex]::Matches($cxml, '\$\{([^}]+)\}')) {
            $tokenCounts[$m.Groups[1].Value] = 1 + $tokenCounts[$m.Groups[1].Value]
        }
        foreach ($k in ($tokenCounts.Keys | Sort-Object)) {
            Write-Host ("  [manifest] {0} x{1}" -f $k, $tokenCounts[$k])
            $manifest.Add(('{0}|{1}|{2}' -f $srcFile.Name, $k, $tokenCounts[$k]))
        }
    } finally { $zin.Dispose() }
}

# --- 寫 manifest ---
[System.IO.File]::WriteAllLines((Join-Path $OutDir 'tokens.txt'), $manifest, (New-Object System.Text.UTF8Encoding($true)))
Write-Host "`n===== 完成 ====="
Write-Host ("manifest: {0} 行；整體狀態: {1}" -f $manifest.Count, $(if ($allOk) { 'OK' } else { '有錯誤' }))
if (-not $allOk) { exit 1 }
