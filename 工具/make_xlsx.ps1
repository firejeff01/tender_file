# ============================================================
# 產生初始 標案資料.xlsx（開發機一次性執行，需要 Excel）
# 三工作表：標案清單（ListObject 表格）、公司資料、說明
# ============================================================
$ErrorActionPreference = 'Stop'
$OutPath = 'C:\WorkSpace\tender_file\標案資料.xlsx'
$TemplateDir = 'C:\WorkSpace\tender_file\範本'

# 從範本資料夾動態取得「產生:xxx」欄名，確保與程式比對一致
$templateNames = Get-ChildItem $TemplateDir -Filter *.odt | ForEach-Object { $_.BaseName }
if ($templateNames.Count -eq 0) { throw "範本資料夾沒有 odt" }

$headers = @('標案名稱', '機關名稱', '案件名稱', '簽署日期') +
           ($templateNames | ForEach-Object { "產生:$_" }) +
           @('備註', '廠商名稱', '負責人姓名', '聯絡電話', '廠商地址')

if (Test-Path $OutPath) { Remove-Item $OutPath -Force }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Add()
    while ($wb.Sheets.Count -gt 1) { $wb.Sheets.Item($wb.Sheets.Count).Delete() }

    # ---------- 工作表 1：標案清單 ----------
    $ws = $wb.Sheets.Item(1)
    $ws.Name = '標案清單'
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $ws.Cells.Item(1, $i + 1).Value2 = $headers[$i]
    }
    $lastCol = $headers.Count
    $colLetter = { param($n) $s=''; while ($n -gt 0) { $r=($n-1)%26; $s=[char](65+$r)+$s; $n=[int](($n-1)/26) }; $s }
    $lastColL = & $colLetter $lastCol

    # 全部資料欄設為文字格式（防：電話前導零遺失、日期被轉序列值）
    $ws.Range("A:$lastColL").NumberFormat = '@'

    # 建立表格（含一列空白資料列，讓驗證與格式隨表格自動延伸）
    $tableRange = $ws.Range("A1:$($lastColL)2")
    $table = $ws.ListObjects.Add(1, $tableRange, $null, 1)   # xlSrcRange, xlYes(有表頭)
    $table.Name = '標案清單表'
    $table.TableStyle = 'TableStyleMedium2'

    # 產生:xxx 欄加上 是/否 下拉驗證（套在資料列範圍）
    for ($i = 0; $i -lt $headers.Count; $i++) {
        if ($headers[$i] -like '產生:*') {
            $cl = & $colLetter ($i + 1)
            $rng = $ws.Range("$($cl)2:$($cl)2")
            $rng.Validation.Delete()
            # xlValidateList=3, xlValidAlertStop=1, xlBetween=1
            $rng.Validation.Add(3, 1, 1, '是,否') | Out-Null
            $rng.Validation.InputMessage = '選「是」或「否」；留空視同「是」'
            $rng.Value2 = '是'
        }
    }

    # 欄寬與外觀
    $ws.Columns.Item(1).ColumnWidth = 42   # 標案名稱
    $ws.Columns.Item(2).ColumnWidth = 16   # 機關名稱
    $ws.Columns.Item(3).ColumnWidth = 42   # 案件名稱
    $ws.Columns.Item(4).ColumnWidth = 14   # 簽署日期
    for ($i = 5; $i -le 4 + $templateNames.Count; $i++) { $ws.Columns.Item($i).ColumnWidth = 13 }
    $ws.Columns.Item(5 + $templateNames.Count).ColumnWidth = 20     # 備註
    for ($i = 6 + $templateNames.Count; $i -le $lastCol; $i++) { $ws.Columns.Item($i).ColumnWidth = 22 }
    $hdr = $ws.Range("A1:$($lastColL)1")
    $hdr.WrapText = $true
    $ws.Rows.Item(1).RowHeight = 42

    # 凍結首列
    $ws.Activate()
    $excel.ActiveWindow.SplitRow = 1
    $excel.ActiveWindow.FreezePanes = $true

    # ---------- 工作表 2：公司資料 ----------
    $ws2 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $ws)
    $ws2.Name = '公司資料'
    $rows2 = @(
        @('參數名稱', '值', '說明'),
        @('廠商名稱', '', '必填。公司登記全名，例：恩潔○○股份有限公司'),
        @('負責人姓名', '', '必填。負責人姓名'),
        @('聯絡電話', '', '必填。例：(07)123-4567'),
        @('廠商地址', '', '必填。完整地址，例：高雄市○○區○○路1號2樓'),
        @('替換後文字顏色', '黑色', '產出文件中被替換文字的顏色：黑色（正式）或 紅色（校對用）')
    )
    for ($r = 0; $r -lt $rows2.Count; $r++) {
        for ($c = 0; $c -lt 3; $c++) { $ws2.Cells.Item($r + 1, $c + 1).Value2 = $rows2[$r][$c] }
    }
    $ws2.Range('B:B').NumberFormat = '@'
    $ws2.Range('A1:C1').Font.Bold = $true
    $colorRng = $ws2.Range('B6')
    $colorRng.Validation.Delete()
    $colorRng.Validation.Add(3, 1, 1, '黑色,紅色') | Out-Null
    $ws2.Columns.Item(1).ColumnWidth = 18
    $ws2.Columns.Item(2).ColumnWidth = 36
    $ws2.Columns.Item(3).ColumnWidth = 55

    # ---------- 工作表 3：說明 ----------
    $ws3 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $ws2)
    $ws3.Name = '說明'
    $lines = @(
        '【標案文件產生器 使用步驟】',
        '',
        '1. 第一次使用：先到「公司資料」工作表，把公司名稱、負責人、電話、地址填好（只需填一次）。',
        '2. 每接一個標案：到「標案清單」工作表新增一列：',
        '   ・標案名稱：必填，會成為輸出資料夾的名稱',
        '   ・機關名稱：招標機關，例：花蓮縣政府',
        '   ・案件名稱：文件中引用的案名；留空會直接使用「標案名稱」',
        '   ・簽署日期：民國格式，例：115年1月1日 或 115/1/1；留空會用「產生當天」的日期',
        '   ・產生:○○○ 欄：選「是」或「否」決定要產生哪些文件；留空視同「是」',
        '3. 存檔並關閉 Excel。',
        '4. 雙擊「標案文件產生器.exe」，勾選要產生的標案，按「產生文件」。',
        '5. 產出的檔案在「輸出\標案名稱\」資料夾內。',
        '',
        '【注意事項】',
        '・「範本」資料夾內的檔案請勿開啟後重新儲存，會破壞參數標記；需要調整範本請聯絡資訊人員。',
        '・同名標案資料夾已存在時，預設會略過不重新產生；要重產請在程式中勾選「重新產生」。',
        '・若某一列的公司資料與預設不同（例如用不同公司投標），可在「標案清單」右側的',
        '  廠商名稱／負責人姓名／聯絡電話／廠商地址 欄位個別填寫，該列會優先使用填入的值。'
    )
    for ($r = 0; $r -lt $lines.Count; $r++) { $ws3.Cells.Item($r + 1, 1).Value2 = $lines[$r] }
    $ws3.Columns.Item(1).ColumnWidth = 110
    $ws3.Cells.Item(1, 1).Font.Bold = $true

    $ws.Activate()   # 開檔時停在標案清單
    $wb.SaveAs($OutPath, 51)   # xlOpenXMLWorkbook
    $wb.Close($false)
    Write-Host "已產生 $OutPath"
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}