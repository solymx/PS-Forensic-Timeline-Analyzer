<#
.SYNOPSIS
    資安鑑識 - 全功能分析工具 (GUI版) v8.0
.DESCRIPTION
    1. 修復：[重大] 修正停止按鈕無反應的問題 (改用正確的計數器刷新 UI)。
    2. 修復：預先掃描擁有者現在正確套用排除清單。
    3. 功能：時間軸、MOTW 解析、Owner 分析、排除清單。
.AUTHOR
    Gemini AI Partner
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Collections

# --- 全域變數 ---
$script:isCancelled = $false
$global:exclusionList = @() 

# --- 設定視窗 ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "資安鑑識 - 時間軸、MOTW 與 排除過濾工具 v8.0 (Stop Fixed)"
$form.Size = New-Object System.Drawing.Size(1200, 800)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 9)

# --- 區塊 1: 路徑設定 ---
$grpPath = New-Object System.Windows.Forms.GroupBox
$grpPath.Text = "1. 搜尋目標 (Target)"
$grpPath.Location = New-Object System.Drawing.Point(15, 10)
$grpPath.Size = New-Object System.Drawing.Size(1150, 60)
$form.Controls.Add($grpPath)

$txtPath = New-Object System.Windows.Forms.TextBox
$txtPath.Location = New-Object System.Drawing.Point(15, 22)
$txtPath.Size = New-Object System.Drawing.Size(1000, 25)
$txtPath.Text = "C:\" 
$grpPath.Controls.Add($txtPath)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "瀏覽..."
$btnBrowse.Location = New-Object System.Drawing.Point(1030, 20)
$btnBrowse.Add_Click({
    $d = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($d.ShowDialog() -eq 'OK') { $txtPath.Text = $d.SelectedPath }
})
$grpPath.Controls.Add($btnBrowse)

# --- 區塊 2: 排除清單 ---
$grpExclude = New-Object System.Windows.Forms.GroupBox
$grpExclude.Text = "2. 排除設定 (Exclusions) - 加速掃描用"
$grpExclude.Location = New-Object System.Drawing.Point(15, 80)
$grpExclude.Size = New-Object System.Drawing.Size(1150, 60)
$form.Controls.Add($grpExclude)

$btnLoadExclude = New-Object System.Windows.Forms.Button
$btnLoadExclude.Text = "載入排除清單 (.txt)..."
$btnLoadExclude.Location = New-Object System.Drawing.Point(15, 20)
$btnLoadExclude.Size = New-Object System.Drawing.Size(160, 30)
$btnLoadExclude.BackColor = "WhiteSmoke"
$grpExclude.Controls.Add($btnLoadExclude)

$lblExcludeStatus = New-Object System.Windows.Forms.Label
$lblExcludeStatus.Text = "目前尚未載入排除規則 (將掃描所有檔案)"
$lblExcludeStatus.Location = New-Object System.Drawing.Point(190, 25)
$lblExcludeStatus.AutoSize = $true
$lblExcludeStatus.ForeColor = "DarkRed"
$grpExclude.Controls.Add($lblExcludeStatus)

# --- 區塊 3: 擁有者分析 ---
$grpOwner = New-Object System.Windows.Forms.GroupBox
$grpOwner.Text = "3. 擁有者分析 (Owner Analysis)"
$grpOwner.Location = New-Object System.Drawing.Point(15, 150)
$grpOwner.Size = New-Object System.Drawing.Size(1150, 60)
$form.Controls.Add($grpOwner)

$btnScanOwners = New-Object System.Windows.Forms.Button
$btnScanOwners.Text = "預先掃描擁有者"
$btnScanOwners.Location = New-Object System.Drawing.Point(15, 20)
$btnScanOwners.Size = New-Object System.Drawing.Size(160, 30)
$btnScanOwners.BackColor = "LightGoldenrodYellow"
$grpOwner.Controls.Add($btnScanOwners)

$lblOwnerFilter = New-Object System.Windows.Forms.Label
$lblOwnerFilter.Text = "過濾擁有者:"
$lblOwnerFilter.Location = New-Object System.Drawing.Point(190, 25)
$lblOwnerFilter.AutoSize = $true
$grpOwner.Controls.Add($lblOwnerFilter)

$cmbOwnerFilter = New-Object System.Windows.Forms.ComboBox
$cmbOwnerFilter.Location = New-Object System.Drawing.Point(280, 22)
$cmbOwnerFilter.Width = 350
$cmbOwnerFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbOwnerFilter.Items.Add("--- 所有擁有者 (All) ---")
$cmbOwnerFilter.SelectedIndex = 0
$grpOwner.Controls.Add($cmbOwnerFilter)

# --- 區塊 4: 時間與 MOTW ---
$grpTime = New-Object System.Windows.Forms.GroupBox
$grpTime.Text = "4. 時間與特徵篩選"
$grpTime.Location = New-Object System.Drawing.Point(15, 220)
$grpTime.Size = New-Object System.Drawing.Size(1150, 60)
$form.Controls.Add($grpTime)

$cmbTimeType = New-Object System.Windows.Forms.ComboBox
$cmbTimeType.Location = New-Object System.Drawing.Point(15, 25)
$cmbTimeType.Width = 140
$cmbTimeType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbTimeType.Items.Add("LastWriteTime")
$cmbTimeType.Items.Add("CreationTime")
$cmbTimeType.Items.Add("LastAccessTime")
$cmbTimeType.SelectedIndex = 0 
$grpTime.Controls.Add($cmbTimeType)

$dtpStart = New-Object System.Windows.Forms.DateTimePicker
$dtpStart.Location = New-Object System.Drawing.Point(170, 25)
$dtpStart.Width = 150
$dtpStart.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$dtpStart.CustomFormat = "yyyy/MM/dd HH:mm:ss"
$dtpStart.Value = (Get-Date).AddDays(-7)
$grpTime.Controls.Add($dtpStart)

$lblTo = New-Object System.Windows.Forms.Label
$lblTo.Text = "~"
$lblTo.Location = New-Object System.Drawing.Point(330, 28)
$lblTo.AutoSize = $true
$grpTime.Controls.Add($lblTo)

$dtpEnd = New-Object System.Windows.Forms.DateTimePicker
$dtpEnd.Location = New-Object System.Drawing.Point(350, 25)
$dtpEnd.Width = 150
$dtpEnd.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$dtpEnd.CustomFormat = "yyyy/MM/dd HH:mm:ss"
$dtpEnd.Value = Get-Date
$grpTime.Controls.Add($dtpEnd)

$chkMOTW = New-Object System.Windows.Forms.CheckBox
$chkMOTW.Text = "深度解析 MOTW (Zone+URLs)"
$chkMOTW.Location = New-Object System.Drawing.Point(530, 27)
$chkMOTW.AutoSize = $true
$chkMOTW.Checked = $true 
$chkMOTW.ForeColor = "DarkBlue"
$grpTime.Controls.Add($chkMOTW)

# --- 區塊 5: 按鈕與結果 ---
$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text = "開始鑑識搜尋"
$btnSearch.Location = New-Object System.Drawing.Point(15, 290)
$btnSearch.Size = New-Object System.Drawing.Size(120, 35)
$btnSearch.BackColor = "LightBlue"
$form.Controls.Add($btnSearch)

$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Text = "停止"
$btnStop.Location = New-Object System.Drawing.Point(145, 290)
$btnStop.Size = New-Object System.Drawing.Size(100, 35)
$btnStop.BackColor = "LightPink"
$btnStop.Enabled = $false
$form.Controls.Add($btnStop)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "匯出 CSV"
$btnExport.Location = New-Object System.Drawing.Point(255, 290)
$btnExport.Size = New-Object System.Drawing.Size(120, 35)
$btnExport.Enabled = $false
$form.Controls.Add($btnExport)

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "就緒。"
$statusStrip.Items.Add($statusLabel)
$form.Controls.Add($statusStrip)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(15, 340)
$grid.Size = New-Object System.Drawing.Size(1150, 400)
$grid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$grid.AllowUserToAddRows = $false
$grid.ReadOnly = $true
$grid.SelectionMode = "FullRowSelect"
$form.Controls.Add($grid)

# 定義欄位
$grid.Columns.Add("MatchTime", "過濾時間點") | Out-Null; $grid.Columns[0].Width = 140
$grid.Columns.Add("MatchType", "類型") | Out-Null; $grid.Columns[1].Width = 80
$grid.Columns.Add("ZoneId", "Zone") | Out-Null; $grid.Columns[2].Width = 50
$grid.Columns.Add("HostUrl", "來源 URL (Host)") | Out-Null; $grid.Columns[3].Width = 250
$grid.Columns.Add("ReferrerUrl", "參照 URL (Referrer)") | Out-Null; $grid.Columns[4].Width = 200
$grid.Columns.Add("Owner", "擁有者") | Out-Null; $grid.Columns[5].Width = 150
$grid.Columns.Add("FullName", "完整路徑") | Out-Null; $grid.Columns[6].Width = 300
$grid.Columns.Add("Length", "KB") | Out-Null; $grid.Columns[7].Width = 60

# --- 輔助函式: UI 切換 ---
function Set-UIState([bool]$isScanning) {
    if ($isScanning) {
        $btnSearch.Enabled = $false; $btnScanOwners.Enabled = $false
        $btnLoadExclude.Enabled = $false; $btnExport.Enabled = $false
        $btnStop.Enabled = $true; $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $script:isCancelled = $false
    } else {
        $btnSearch.Enabled = $true; $btnScanOwners.Enabled = $true
        $btnLoadExclude.Enabled = $true; $btnStop.Enabled = $false
        if ($grid.Rows.Count -gt 0) { $btnExport.Enabled = $true }
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

# --- 輔助函式: 檢查是否需排除 ---
function Is-Excluded {
    param($filePath)
    if ($global:exclusionList.Count -eq 0) { return $false }
    foreach ($exPath in $global:exclusionList) {
        if ($filePath.StartsWith($exPath, [StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

# --- 邏輯: 載入排除清單 ---
$btnLoadExclude.Add_Click({
    $openDlg = New-Object System.Windows.Forms.OpenFileDialog
    $openDlg.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    if ($openDlg.ShowDialog() -eq 'OK') {
        try {
            $rawContent = Get-Content $openDlg.FileName
            $global:exclusionList = @()
            foreach ($line in $rawContent) {
                $tLine = $line.Trim()
                if ($tLine -ne "" -and -not $tLine.StartsWith("#")) {
                    $global:exclusionList += $tLine
                }
            }
            $lblExcludeStatus.Text = "已載入 $($global:exclusionList.Count) 條排除規則。"
            $lblExcludeStatus.ForeColor = "Green"
            [System.Windows.Forms.MessageBox]::Show("排除清單載入成功！", "成功", "OK", "Information")
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("讀取檔案失敗: $($_.Exception.Message)", "錯誤", "OK", "Error")
        }
    }
})

# --- 輔助: MOTW 解析 ---
function Get-MOTWDetails {
    param($filePath)
    $result = [PSCustomObject]@{ ZoneId="-"; HostUrl=""; ReferrerUrl="" }
    $adsPath = "$filePath`:Zone.Identifier"
    if (Test-Path -LiteralPath $adsPath) {
        try {
            $content = Get-Content -LiteralPath $filePath -Stream Zone.Identifier -Raw -ErrorAction SilentlyContinue
            if ($content -match "ZoneId=(\d)") { $result.ZoneId = $Matches[1] }
            if ($content -match "HostUrl=(.*)") { $result.HostUrl = $Matches[1].Trim() }
            if ($content -match "ReferrerUrl=(.*)") { $result.ReferrerUrl = $Matches[1].Trim() }
        } catch { $result.ZoneId = "Error" }
    }
    return $result
}

$btnStop.Add_Click({ 
    $script:isCancelled = $true
    $statusLabel.Text = "正在嘗試停止... 請稍候..."
})

# --- 邏輯: 掃描擁有者 (修復版) ---
$btnScanOwners.Add_Click({
    $path = $txtPath.Text
    if (-not (Test-Path $path)) { [System.Windows.Forms.MessageBox]::Show("路徑不存在"); return }
    
    Set-UIState -isScanning $true
    $statusLabel.Text = "正在掃描擁有者清單..."
    $cmbOwnerFilter.Items.Clear(); $cmbOwnerFilter.Items.Add("--- 所有擁有者 (All) ---"); $cmbOwnerFilter.SelectedIndex = 0
    [System.Windows.Forms.Application]::DoEvents()
    
    $uniqueOwners = New-Object System.Collections.Generic.HashSet[string]
    $counter = 0 # 計數器

    try {
        Get-ChildItem -Path $path -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
            $counter++
            
            # [修復] 使用計數器而非檔案大小，每 20 個檔案檢查一次停止
            if ($counter % 20 -eq 0) { 
                [System.Windows.Forms.Application]::DoEvents() 
                if ($script:isCancelled) { throw "USER_CANCEL" }
            }

            if (Is-Excluded -filePath $_.FullName) { return }

            try { 
                $acl = $_.GetAccessControl()
                if ($acl -and $acl.Owner) { 
                    $uniqueOwners.Add($acl.Owner) | Out-Null 
                } 
            } catch {}
            
            if ($counter % 50 -eq 0) { 
                $statusLabel.Text = "已掃描 $counter 個檔案，發現 $($uniqueOwners.Count) 個擁有者..." 
            }
        }
        $statusLabel.Text = "掃描完成。共發現 $($uniqueOwners.Count) 個擁有者。"
    }
    catch { 
        if ($_.Exception.Message -eq "USER_CANCEL") { $statusLabel.Text = "使用者已中斷掃描。" }
        else { $statusLabel.Text = "錯誤: $($_.Exception.Message)" }
    }
    finally {
        $sortedOwners = $uniqueOwners | Sort-Object
        foreach ($o in $sortedOwners) { $cmbOwnerFilter.Items.Add($o) | Out-Null }
        Set-UIState -isScanning $false
    }
})

# --- 邏輯: 搜尋主程式 (修復版) ---
$global:searchResults = @()

$btnSearch.Add_Click({
    $path = $txtPath.Text
    $start = $dtpStart.Value
    $end = $dtpEnd.Value
    $targetOwner = $cmbOwnerFilter.SelectedItem.ToString()
    $isAllOwners = ($targetOwner -eq "--- 所有擁有者 (All) ---")
    $needMOTW = $chkMOTW.Checked

    $timeProperty = ""; $timeLabel = ""
    switch ($cmbTimeType.SelectedIndex) {
        0 { $timeProperty = "LastWriteTime"; $timeLabel = "Modified" }
        1 { $timeProperty = "CreationTime";  $timeLabel = "Created" }
        2 { $timeProperty = "LastAccessTime"; $timeLabel = "Accessed" }
    }

    if (-not (Test-Path $path)) { [System.Windows.Forms.MessageBox]::Show("路徑錯誤"); return }

    $grid.Rows.Clear(); $global:searchResults = @(); Set-UIState -isScanning $true
    $statusLabel.Text = "正在搜尋..."
    [System.Windows.Forms.Application]::DoEvents()

    $counter = 0 # 計數器

    try {
        Get-ChildItem -Path $path -Recurse -Force -File -ErrorAction SilentlyContinue | 
        ForEach-Object {
            $counter++

            # [修復] 使用計數器刷新 UI 和檢查停止
            if ($counter % 20 -eq 0) { 
                [System.Windows.Forms.Application]::DoEvents() 
                if ($script:isCancelled) { throw "USER_CANCEL" }
            }

            $file = $_
            
            # 排除檢查
            if (Is-Excluded -filePath $file.FullName) { return }

            $checkTime = $file.$timeProperty

            if ($checkTime -ge $start -and $checkTime -le $end) {
                
                $fileOwner = "Unknown"
                try { $acl = $file.GetAccessControl(); if ($acl) { $fileOwner = $acl.Owner } } catch { $fileOwner = "Access Denied" }

                if ($isAllOwners -or ($fileOwner -eq $targetOwner)) {
                    
                    $motwData = [PSCustomObject]@{ ZoneId="-"; HostUrl=""; ReferrerUrl="" }
                    if ($needMOTW) { $motwData = Get-MOTWDetails -filePath $file.FullName }

                    $sizeKB = [math]::Round($file.Length / 1KB, 2)
                    $timeStr = $checkTime.ToString("yyyy/MM/dd HH:mm:ss")

                    $grid.Rows.Add($timeStr, $timeLabel, $motwData.ZoneId, $motwData.HostUrl, $motwData.ReferrerUrl, $fileOwner, $file.FullName, $sizeKB) | Out-Null
                    
                    $global:searchResults += [PSCustomObject]@{
                        MatchTime=$timeStr; MatchType=$timeLabel; ZoneId=$motwData.ZoneId;
                        HostUrl=$motwData.HostUrl; ReferrerUrl=$motwData.ReferrerUrl; Owner=$fileOwner;
                        FullName=$file.FullName; SizeKB=$sizeKB
                    }
                    
                    if ($grid.Rows.Count % 20 -eq 0) { $statusLabel.Text = "找到 $($grid.Rows.Count) 個檔案..." }
                }
            }
        }
        $statusLabel.Text = "搜尋完成。共 $($grid.Rows.Count) 個檔案。"
    }
    catch {
        if ($_.Exception.Message -eq "USER_CANCEL") { 
            $statusLabel.Text = "搜尋已中斷。" 
            [System.Windows.Forms.MessageBox]::Show("您已成功中斷搜尋。", "資訊", "OK", "Warning")
        }
        else { $statusLabel.Text = "錯誤: $($_.Exception.Message)" }
    }
    finally { Set-UIState -isScanning $false }
})

# --- 匯出 CSV ---
$btnExport.Add_Click({
    if ($global:searchResults.Count -eq 0) { return }
    $s = New-Object System.Windows.Forms.SaveFileDialog
    $s.Filter = "CSV Files|*.csv"; $s.FileName = "Forensic_Result_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
    if ($s.ShowDialog() -eq 'OK') {
        $global:searchResults | Export-Csv -Path $s.FileName -NoTypeInformation -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("匯出完成！", "資訊", "OK", "Information")
    }
})

$form.ShowDialog() | Out-Null
$form.Dispose()
