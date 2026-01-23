# ==========================================
# 0. Global Setup（全域環境設定)
# ==========================================
# 設定錯誤處理模式：遇到任何錯誤直接停止，方便我們攔截並報警，而不是讓它默默出錯。
$ErrorActionPreference = 'Stop' 
# 強制 Console 輸出使用 UTF-8，避免中文亂碼
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# 強制清除可能殘留的全域變數，解決重複執行腳本時出現舊紀錄的問題
Remove-Variable SessionLog -Scope Global -ErrorAction SilentlyContinue
Remove-Variable LastReportLogIndex -Scope Global -ErrorAction SilentlyContinue
Remove-Variable LastBitmapCache -Scope Global -ErrorAction SilentlyContinue

try {
    # --- 特殊符號定義 (使用 ASCII/Unicode 編碼避免亂碼) ---
    $CR = [char]13  # 歸位字元 (回到行首，用於進度條動畫)
    $LF = [char]10  # 換行字元

    $Icon_Warn   = [string][char]0x26A0 + [char]0xFE0F # ⚠️
    $Icon_Check  = [char]0x2705 # ✅
    $Icon_Cross  = [char]0x274C # ❌
    $Icon_Heart  = [string][char]0xD83E + [char]0xDEC0 # 🫀
    $Icon_Bullet = [char]0x2022 # •
    $Icon_Start  = [string][char]0x26A1 + [char]0x26BD # ⚡⚽
    $Icon_Stop   = [char]0x23F9 # ⏹️

    # --- 中文訊息設定 (集中管理，方便修改) ---
    $Msg_Title_Start    = '看門狗 v1.0.3 已啟動'
    $Msg_Reason_Start   = '啟動通知'
    $Msg_Title_Stop     = '監控已手動停止'
    $Msg_Reason_Stop    = '使用者主動結束監控'
    
    # 狀態顯示文字
    $Msg_Game_Run       = '勝利之路：運作中'
    $Msg_Game_NoResp    = '勝利之路：無回應'
    $Msg_Game_Lost      = '勝利之路：消失'
    $Msg_KTK_Run        = 'KeyToKey：運作中'
    $Msg_KTK_Err        = 'KeyToKey：異常'
    
    # 錯誤與警告訊息
    $Msg_Warn_NoResp    = '警告：程式無回應'
    $Msg_Warn_Freeze    = '警告：畫面凍結 (相似度 '
    $Msg_Err_Freeze     = '畫面凍結 (死機)'
    $Msg_Err_NoResp     = '程式無回應 (卡死)'
    $Msg_Err_Crash      = '程式崩潰 (消失)'
    $Msg_Err_Sys        = '偵測到系統錯誤：ID'
    $Msg_Err_Reason     = '系統嚴重錯誤 (ID:'
    $Msg_Prot_Trig      = '觸發保護：'
    
    # Discord 回報標題與內容
    $Msg_Discord_Title  = '嚴重異常終止'
    $Msg_Discord_Title_W= '異常徵兆警告'
    $Msg_Discord_HB     = '看門狗定期報告'
    $Msg_Discord_HBTxt  = '定期健康報告'
    $Msg_Discord_SysOK  = '勝利之路：🟢    |   KeyToKey：🟢 '
    $Msg_Discord_Log    = '近期紀錄：'
    
    # 其他介面訊息
    $Msg_Shutdown       = '60秒後關機...'
    $Msg_GUI_Title      = '⚠️ 掛機失敗——關機預警'
    $Msg_GUI_Cancel     = '🚫 取消關機'
    $Msg_Stop_Monitor   = '監控已停止。按 Enter 離開視窗...'
    $Msg_Status_OK      = '掛機運作中'
    $Msg_Sent_Report    = '已發送定期 Discord 報告'
    $Msg_KTK_Restart    = 'KeyToKey 重啟中...'
    $Msg_Wait_Load      = '等待 35 秒載入...'
    $Msg_Send_Key       = '發送按鍵'
    $Msg_Recovered      = '復原完畢'
    $Msg_Footer_Base    = 'Watchdog v1.0.3'
    $Msg_Ask_Webhook    = '[設定] 初次執行，請輸入 Discord Webhook 網址 (輸入完畢按 Enter):'
    $Msg_Webhook_Saved  = '網址已儲存至 webhook.txt，下次將自動讀取。'

    $ScriptStartTime = Get-Date
    $DiscordUserID   = '649980145020436497' 

    # ==========================================
    # 1. User Settings（使用者設定區）
    # ==========================================
    # 請確保此路徑正確，否則程式會發出警告
    $KeyToKeyPath = 'D:\Users\user\Downloads\KeyToKey\KeyToKey.exe'
    $ScreenshotDir = "D:\Users\user\Desktop\勝利之路看門狗"
    $LogSavePath = $env:USERPROFILE + '\Desktop\Watchdog_Log_Latest.txt'

    # Webhook 讀取邏輯：優先讀取同目錄下的 webhook.txt
    $ScriptPath = $MyInvocation.MyCommand.Path
    $ScriptDir  = Split-Path $ScriptPath -Parent
    $WebhookFile = Join-Path $ScriptDir 'webhook.txt'
    if (Test-Path $WebhookFile) {
        $DiscordWebhookUrl = (Get-Content $WebhookFile -Raw).Trim()
    } else { $DiscordWebhookUrl = '' }

    # 監控參數
    $LoopIntervalSeconds = 75  # 每次檢測的間隔秒數 (降為 75 秒，加上處理時間後，Log 紀錄的間隔會接近 90 秒)
    $FreezeThreshold = 3       # 連續畫面凍結幾次才判定為當機
    $NoResponseThreshold = 3   # 連續無回應幾次才判定為卡死
    $FreezeSimilarity = 98.5   # 判定畫面凍結的相似度門檻 (%)

    # 初始化全域變數
    $Global:SessionLog = @()
    $Global:LastReportLogIndex = 0 # 紀錄上次回報到 Log 的哪一行
    $Global:LastHeartbeatTime = Get-Date
    $Global:HeartbeatInterval = 5 # 分鐘
    $Global:LastBitmapCache = $null # 初始化上一張畫面的緩存 (用於凍結對比)

    # ==========================================
    # 2. System Core（系統核心與 Windows API）
    # ==========================================
    # 設定 Console 緩衝區大小，避免文字太長被截斷
    try {
        $PSWindow = (Get-Host).UI.RawUI
        $BufferSize = $PSWindow.BufferSize; $BufferSize.Width = 120; $PSWindow.BufferSize = $BufferSize
        $WindowSize = $PSWindow.WindowSize; $WindowSize.Width = 120; $PSWindow.WindowSize = $WindowSize
    } catch {}

    # 確保截圖目錄存在
    if (!(Test-Path $ScreenshotDir)) { New-Item -ItemType Directory -Path $ScreenshotDir | Out-Null }
    
    # 載入 .NET 繪圖與表單組件
    Add-Type -AssemblyName System.Drawing, System.Net.Http, System.Windows.Forms

    # C# 核心代碼：整合了視窗控制與 Steam 攔截功能 (SteamBuster)
    # 使用底層 EnumWindows API 來進行「地毯式搜索」，解決 Get-Process 找不到子視窗的問題
    $Win32Code = @'
    using System;
    using System.Collections.Generic;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Diagnostics;

    public class Win32Tools {
        [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
        
        // 硬體級按鍵模擬 API
        [DllImport("user32.dll")] public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
        public const int KEYEVENTF_KEYUP = 0x0002;
        public const int SW_RESTORE = 9;
    }

    public class SteamBuster {
        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);
        [DllImport("user32.dll")] private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        private const uint WM_CLOSE = 0x0010;

        public class WindowInfo {
            public IntPtr Handle;
            public string Title;
            public string ProcessName;
        }

        // 搜尋所有屬於 Steam 相關處理程序 (Process) 的視窗
        public static List<WindowInfo> FindSteamWindows() {
            var list = new List<WindowInfo>();
            EnumWindows((hWnd, lParam) => {
                if (IsWindowVisible(hWnd)) {
                    int pid;
                    GetWindowThreadProcessId(hWnd, out pid);
                    try {
                        Process p = Process.GetProcessById(pid);
                        string pName = p.ProcessName.ToLower();
                        if (pName == "steam" || pName == "steamwebhelper") {
                            StringBuilder sb = new StringBuilder(256);
                            GetWindowText(hWnd, sb, 256);
                            string title = sb.ToString();
                            if (!string.IsNullOrEmpty(title)) { 
                                 list.Add(new WindowInfo { Handle = hWnd, Title = title, ProcessName = pName });
                            }
                        }
                    } catch {}
                }
                return true;
            }, IntPtr.Zero);
            return list;
        }

        public static void CloseWindow(IntPtr hWnd) { PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero); }
    }
'@
    # [Fix] 檢查 Win32Tools 是否已定義，防止重複執行腳本時發生崩潰
    if (-not ("Win32Tools" -as [type])) {
        Add-Type -TypeDefinition $Win32Code
    }

    try { [Console]::CursorVisible = $false } catch {}
    try { [Win32Tools]::SetProcessDPIAware() | Out-Null } catch {}

    # ==========================================
    # 3. Helpers（輔助功能函式庫）
    # ==========================================

    # -----------------------------------------------------------
    # 函式：Write-Log
    # 功能：將訊息寫入控制台，並同步記錄到記憶體中的全域日誌變數
    # -----------------------------------------------------------
    function Write-Log {
        param($Message, $Color='White', $ForceNewLine=$false)
        $Time = Get-Date -Format 'HH:mm:ss'
        $LogLine = '[' + $Time + '] ' + $Message
        
        # 確保倒數計時被中斷時能正確換行顯示
        if ($ForceNewLine) { Write-Host '' } 
        Write-Host ($LogLine + '          ') -ForegroundColor $Color
        $Global:SessionLog += $LogLine
        
        # [記憶體保護] 限制日誌緩存最多 1000 行，避免長時間掛機佔用過多記憶體
        if ($Global:SessionLog.Count -gt 1000) { 
            $Global:SessionLog = $Global:SessionLog[-1000..-1] 
            $Global:LastReportLogIndex = [Math]::Max(0, $Global:LastReportLogIndex - ($Global:SessionLog.Count - 1000))
        }
    }

    # -----------------------------------------------------------
    # 函式：Ensure-Game-TopMost
    # 功能：檢查遊戲視窗，若被縮小則還原，並嘗試將其設為前景
    # -----------------------------------------------------------
    function Ensure-Game-TopMost {
        $GameProc = Get-Process -Name 'nie' -ErrorAction SilentlyContinue
        if ($GameProc) {
            $Handle = $GameProc.MainWindowHandle
            if ($Handle -ne [IntPtr]::Zero) {
                if ([Win32Tools]::IsIconic($Handle)) {
                    [Win32Tools]::ShowWindow($Handle, [Win32Tools]::SW_RESTORE) | Out-Null
                }
                [Win32Tools]::SetForegroundWindow($Handle) | Out-Null
            }
        }
    }

    # -----------------------------------------------------------
    # 函式：Send-Key-Native
    # 功能：使用 keybd_event 模擬實體鍵盤訊號 (用於重啟 KTK)
    # -----------------------------------------------------------
    function Send-Key-Native ($KeyName) {
        try {
            # 1. 嘗試解析按鍵代碼 (例如 "F7" -> 118)
            $VK = [System.Windows.Forms.Keys]::Parse([System.Windows.Forms.Keys], $KeyName)
            $KeyCode = [byte][int]$VK
            
            # 2. 使用硬體訊號模擬：按下 -> 等待 -> 放開
            [Win32Tools]::keybd_event($KeyCode, 0, 0, [UIntPtr]::Zero)
            Start-Sleep -Milliseconds 100
            [Win32Tools]::keybd_event($KeyCode, 0, [Win32Tools]::KEYEVENTF_KEYUP, [UIntPtr]::Zero)
            
            return $true
        } catch {
            # 3. 如果解析失敗，退回舊版 SendKeys
            try {
                $KeyStr = '{' + $KeyName + '}'
                [System.Windows.Forms.SendKeys]::SendWait($KeyStr)
                return $true
            } catch { return $false }
        }
    }
    # -----------------------------------------------------------
    # 函式：Show-Crash-Warning-GUI
    # 功能：顯示警告視窗，倒數 60 秒後自動關機
    # -----------------------------------------------------------
    function Show-Crash-Warning-GUI {
        param([string]$Reason)

        # --- 1. 定義配色 ---
        $Color_Bg = [System.Drawing.Color]::FromArgb(30, 30, 30)        # 深灰背景
        $Color_Accent = [System.Drawing.Color]::FromArgb(255, 60, 60)   # 亮警報紅
        $Color_TextPri = [System.Drawing.Color]::White               # 純白文字
        $Color_TextSec = [System.Drawing.Color]::FromArgb(200, 200, 200)# 淺灰文字
        
        # --- 2. 表單設定 ---
        $Form = New-Object System.Windows.Forms.Form
        $Form.Size = New-Object System.Drawing.Size(600, 380)
        $Form.StartPosition = 'CenterScreen'
        $Form.TopMost = $true
        $Form.FormBorderStyle = 'None' # 無邊框
        $Form.BackColor = $Color_Accent 
        $Form.Padding = New-Object System.Windows.Forms.Padding(4) # 紅色邊框

        $MainPanel = New-Object System.Windows.Forms.Panel
        $MainPanel.Dock = 'Fill'; $MainPanel.BackColor = $Color_Bg; $Form.Controls.Add($MainPanel)

        # --- 3. UI 元件 (使用自動置中) ---
        
        # [標題]
        $LblTitle = New-Object System.Windows.Forms.Label
        # [Fix] 這裡改用全域變數 $Msg_GUI_Title 以便支援設定修改
        $LblTitle.Text = $script:Msg_GUI_Title 
        $LblTitle.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 20, [System.Drawing.FontStyle]::Bold)
        $LblTitle.ForeColor = $Color_Accent; $LblTitle.AutoSize = $false
        $LblTitle.Size = New-Object System.Drawing.Size(592, 50); $LblTitle.Location = New-Object System.Drawing.Point(0, 30)
        $LblTitle.TextAlign = 'MiddleCenter'; $MainPanel.Controls.Add($LblTitle)

        # [原因]
        $LblReason = New-Object System.Windows.Forms.Label
        $DispReason = if ($Reason.Length -gt 45) { $Reason.Substring(0, 42) + "..." } else { $Reason }
        $LblReason.Text = "$DispReason"
        $LblReason.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 12)
        $LblReason.ForeColor = $Color_TextSec; $LblReason.AutoSize = $false
        $LblReason.Size = New-Object System.Drawing.Size(592, 30); $LblReason.Location = New-Object System.Drawing.Point(0, 80)
        $LblReason.TextAlign = 'MiddleCenter'; $MainPanel.Controls.Add($LblReason)

        # [倒數數字]
        $LblCount = New-Object System.Windows.Forms.Label
        $LblCount.Text = "60"
        $LblCount.Font = New-Object System.Drawing.Font("Arial", 55, [System.Drawing.FontStyle]::Bold)
        $LblCount.ForeColor = $Color_TextPri; $LblCount.AutoSize = $false
        $LblCount.Size = New-Object System.Drawing.Size(592, 100); $LblCount.Location = New-Object System.Drawing.Point(0, 115)
        $LblCount.TextAlign = 'MiddleCenter'; $MainPanel.Controls.Add($LblCount)
        
        # [倒數文字]
        $LblSub = New-Object System.Windows.Forms.Label
        $LblSub.Text = "秒後將執行系統保護關機..."
        $LblSub.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 11)
        $LblSub.ForeColor = $Color_Accent; $LblSub.AutoSize = $false
        $LblSub.Size = New-Object System.Drawing.Size(592, 30); $LblSub.Location = New-Object System.Drawing.Point(0, 215)
        $LblSub.TextAlign = 'TopCenter'; $MainPanel.Controls.Add($LblSub)

        # [取消按鈕]
        $BtnCancel = New-Object System.Windows.Forms.Button
        # [Fix] 這裡改用全域變數 $Msg_GUI_Cancel
        $BtnCancel.Text = $script:Msg_GUI_Cancel
        $BtnCancel.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 16, [System.Drawing.FontStyle]::Bold)
        $BtnCancel.Size = New-Object System.Drawing.Size(260, 60)
        $BtnX = [int]((600 - 260) / 2)
        $BtnCancel.Location = New-Object System.Drawing.Point([int]($BtnX - 4), 270)
        $BtnCancel.BackColor = [System.Drawing.Color]::White; $BtnCancel.ForeColor = [System.Drawing.Color]::Black
        $BtnCancel.FlatStyle = 'Flat'; $BtnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $MainPanel.Controls.Add($BtnCancel)

        # --- 4. Timer 邏輯 ---
        $Timer = New-Object System.Windows.Forms.Timer; $Timer.Interval = 1000; $Script:CountDown = 60
        $Timer.Add_Tick({
            $Script:CountDown--
            $LblCount.Text = "$Script:CountDown"
            if ($Script:CountDown -le 0) { $Timer.Stop(); $Form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $Form.Close() }
        })
        $Timer.Start()
        # 預設聚焦取消按鈕，方便直接按 Enter/空白鍵取消
        $Form.Add_Shown({ $BtnCancel.Focus() })
        
        $Result = $Form.ShowDialog(); $Timer.Stop(); $Form.Dispose(); return $Result
    }

    # -----------------------------------------------------------
    # 函式：Send-Discord-Report
    # 功能：組合訊息、日誌與截圖，發送 Embed 到 Discord
    # -----------------------------------------------------------
    function Send-Discord-Report {
        param([string]$Title, [string]$Reason, [string]$ColorType='Green', [string[]]$ImagePaths=@(), [bool]$IsHeartbeat=$false)
        if ([string]::IsNullOrWhiteSpace($DiscordWebhookUrl) -or $DiscordWebhookUrl -eq 'YOUR_WEBHOOK_HERE') { return }
        
        # 如果不是心跳包 (例如黃色警告)，不要強制換行，避免空行問題
        if (!$IsHeartbeat) { Write-Log 'Uploading Report...' 'Cyan' $false }

        # --- 日誌處理邏輯 ---
        $LogPreviewLines = @()
        if ($IsHeartbeat) {
            # 定期報告：只抓取新紀錄
            $NewCount = $Global:SessionLog.Count
            if ($NewCount -gt $Global:LastReportLogIndex) {
                for ($k = $Global:LastReportLogIndex; $k -lt $NewCount; $k++) { $LogPreviewLines += $Global:SessionLog[$k] }
            }
            $Global:LastReportLogIndex = $NewCount
        } else {
            # 異常報告：抓取最後 15 行
            $MaxLines = 15; $Count = 0
            for ($k = $Global:SessionLog.Count - 1; $k -ge 0; $k--) {
                $LogPreviewLines += $Global:SessionLog[$k]; $Count++
                if ($Count -ge $MaxLines) { break }
            }
            [array]::Reverse($LogPreviewLines)
        }
        $LogPreview = $LogPreviewLines -join $LF
        if ([string]::IsNullOrWhiteSpace($LogPreview)) { $LogPreview = '(無)' }
        $Global:SessionLog | Out-File -FilePath $LogSavePath -Encoding UTF8

        # --- 建立 Embed ---
        $ColorMap = @{ 'Green'=5763719; 'Red'=15548997; 'Yellow'=16705372; 'Blue'=5793266; 'Grey'=9807270 }
        
        $Duration = New-TimeSpan -Start $ScriptStartTime -End (Get-Date)
        $RunTimeStr = "{0:D2}小時{1:D2}分鐘" -f [int][Math]::Floor($Duration.TotalHours), $Duration.Minutes

        $DescHeader = ''
        $MentionContent = ''

        if ($IsHeartbeat) {
            $DescHeader = '**' + $Icon_Check + ' ' + $Msg_Discord_HBTxt + '**' + $LF + $Msg_Discord_SysOK + $LF + 
                          "(每 $Global:HeartbeatInterval 分鐘回報一次)" + $LF + $LF + '⏱️ **已運行時間**' + $LF + $RunTimeStr
        } else {
            # 將「原因」與「掛機時長」分段顯示
            $DescHeader = "**異常原因：**$LF" + $Reason + $LF + $LF + "⏳ **已掛機：**$LF" + $RunTimeStr
            # 紅燈與黃燈都要 @使用者
            if ($ColorType -eq 'Red' -or $ColorType -eq 'Yellow') { $MentionContent = "<@$DiscordUserID>" }
        }

        # 如果是黃色警告，只顯示原因，不顯示日誌預覽區塊 (避免洗版)
        if ($ColorType -ne 'Yellow') {
            $EmbedDesc = $DescHeader + $LF + $LF + '**📋 ' + $Msg_Discord_Log + '**' + $LF + '```' + $LF + $LogPreview + $LF + '```'
        } else { $EmbedDesc = $DescHeader }

        $FooterTxt = $Msg_Footer_Base + ' ' + $Icon_Bullet + ' ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        $Embed = @{ title = $Title; description = $EmbedDesc; color = $ColorMap[$ColorType]; footer = @{ text = $FooterTxt } }
        $Payload = @{ content = $MentionContent; embeds = @($Embed) }
        
        # 轉為 JSON 並確保 UTF-8 編碼
        $JsonPayload = $Payload | ConvertTo-Json -Depth 10 -Compress

        # --- 發送 Multipart HTTP 請求 ---
        $HttpClient = New-Object System.Net.Http.HttpClient
        $Streams = @(); $Form = $null
        try {
            $Form = New-Object System.Net.Http.MultipartFormDataContent
            $Enc = [System.Text.Encoding]::UTF8
            $Form.Add((New-Object System.Net.Http.StringContent($JsonPayload, $Enc, 'application/json')), 'payload_json')

            $ImgIndex = 1
            foreach ($Path in $ImagePaths) {
                if (![string]::IsNullOrEmpty($Path) -and (Test-Path $Path)) {
                    $FS = [System.IO.File]::OpenRead($Path); $Streams += $FS
                    $ImgContent = New-Object System.Net.Http.StreamContent($FS)
                    $ImgContent.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse('image/png')
                    $Form.Add($ImgContent, "file$ImgIndex", [System.IO.Path]::GetFileName($Path))
                    $ImgIndex++
                }
            }
            # 僅在「非心跳」且「非黃色警告」時，才附上完整 Log 文字檔
            if (!$IsHeartbeat -and $ColorType -ne 'Yellow' -and (Test-Path $LogSavePath)) {
                $FS2 = [System.IO.File]::OpenRead($LogSavePath); $Streams += $FS2
                $TxtContent = New-Object System.Net.Http.StreamContent($FS2)
                $TxtContent.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse('text/plain')
                $Form.Add($TxtContent, "file_log", 'Watchdog_Log.txt')
            }
            # 加入 charset=utf-8 確保中文顯示正常
            $HttpClient.PostAsync($DiscordWebhookUrl, $Form).Result | Out-Null
        } catch { Write-Log "Discord 上傳失敗: $_" 'Red' $true } 
        finally {
            foreach ($s in $Streams) { $s.Close(); $s.Dispose() }
            if ($HttpClient) { $HttpClient.Dispose() }
            if ($Form) { $Form.Dispose() }
        }

        # 修復：等待 1 秒確保 HttpClient 徹底釋放檔案，再執行刪除
        Start-Sleep -Seconds 1
        foreach ($Path in $ImagePaths) { if (Test-Path $Path) { try { Remove-Item $Path -Force -ErrorAction SilentlyContinue } catch {} } }
        if (Test-Path $LogSavePath) { try { Remove-Item $LogSavePath -Force -ErrorAction SilentlyContinue } catch {} }
    }

    # -----------------------------------------------------------
    # 函式：Suppress-Steam-Window
    # 功能：偵測並關閉 Steam 干擾視窗，關閉後強制將遊戲視窗置頂
    # -----------------------------------------------------------
    function Suppress-Steam-Window {
        # 使用 C# 核心的 FindSteamWindows 進行全域搜尋
        $Targets = [SteamBuster]::FindSteamWindows()
        $ClosedAny = $false
        foreach ($win in $Targets) {
            # 1. 偵測到干擾：寫 Log + 顯示
            $Msg = "偵測到干擾視窗！標題: [$($win.Title)] (程式: $($win.ProcessName))"
            Write-Log ($Icon_Warn + ' ' + $Msg) 'Yellow' $true
            # 2. 發送警告
            Send-Discord-Report -Title ($Icon_Warn + ' 異常徵兆警告') -Reason "$Msg`n(已執行自動關閉)" -ColorType 'Yellow'
            # 3. 執行關閉
            [SteamBuster]::CloseWindow($win.Handle)
            $ClosedAny = $true
            # 4. 寫 Log
            Write-Log ($Icon_Check + ' 已關閉視窗。') 'Green'
        }
        
        # 如果有關閉過視窗，立刻把焦點搶回遊戲，避免 KTK 按錯地方
        if ($ClosedAny) { Ensure-Game-TopMost }
    }

    # -----------------------------------------------------------
    # 函式：Capture-ScreenBitmap
    # 功能：截取全螢幕畫面並回傳 Bitmap 物件 (不存檔，減少硬碟讀寫)
    # -----------------------------------------------------------
    function Capture-ScreenBitmap {
        try {
            $Bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
            $Bitmap = New-Object System.Drawing.Bitmap $Bounds.Width, $Bounds.Height
            $Graphics = [System.Drawing.Graphics]::FromImage($Bitmap)
            $Graphics.CopyFromScreen($Bounds.Location, [System.Drawing.Point]::Empty, $Bounds.Size)
            $Graphics.Dispose()
            return $Bitmap
        } catch { return $null }
    }

    # -----------------------------------------------------------
    # 函式：Get-PixelsFromBitmap
    # 功能：將高解析度 Bitmap 縮放為 12x12 的縮圖，並提取像素顏色值
    # 目的：透過極低解析度比對，大幅降低 CPU 運算量，同時忽略細微雜訊
    # -----------------------------------------------------------
    function Get-PixelsFromBitmap ($Bitmap) {
        try {
            if (!$Bitmap) { return $null }
            $Small = $Bitmap.GetThumbnailImage(12, 12, $null, [IntPtr]::Zero)
            $Pixels = New-Object 'int[,]' 12, 12
            for ($x=0; $x -lt 12; $x++) { for ($y=0; $y -lt 12; $y++) { $Pixels[$x, $y] = $Small.GetPixel($x, $y).ToArgb() } }
            $Small.Dispose(); return ,$Pixels 
        } catch { return $null }
    }

    # -----------------------------------------------------------
    # 函式：Get-Similarity
    # 功能：比對兩組像素矩陣的相似度 (0~100%)
    # -----------------------------------------------------------
    function Get-Similarity ($PixA, $PixB) {
        if (!$PixA -or !$PixB) { return 0 }
        $Match = 0; $Total = 144
        for ($x=0; $x -lt 12; $x++) { for ($y=0; $y -lt 12; $y++) {
            $valA = $PixA[$x, $y]; $valB = $PixB[$x, $y]
            $R1 = ($valA -shr 16) -band 255; $G1 = ($valA -shr 8) -band 255; $B1 = $valA -band 255
            $R2 = ($valB -shr 16) -band 255; $G2 = ($valB -shr 8) -band 255; $B2 = $valB -band 255
            if ([Math]::Abs($R1 - $R2) -lt 20 -and [Math]::Abs($G1 - $G2) -lt 20 -and [Math]::Abs($B1 - $B2) -lt 20) { $Match++ }
        }}
        return [Math]::Round(($Match / $Total) * 100, 1)
    }

    # -----------------------------------------------------------
    # 函式：Save-BitmapToFile
    # 功能：將記憶體中的 Bitmap 存為實體 PNG 檔案 (用於發送 Discord 附件)
    # -----------------------------------------------------------
    function Save-BitmapToFile ($Bitmap, $Prefix) {
        if (!$Bitmap) { return $null }
        $TimeStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $FName = $Prefix + '_' + $TimeStamp + '.png'
        $Path = Join-Path $ScreenshotDir $FName
        $Bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
        return $Path
    }

    # ==========================================
    # 4. Initialization（初始化流程）
    # ==========================================
    Clear-Host
    try { [Console]::CursorVisible = $true } catch {}
    Write-Host '==========================================' -ForegroundColor Cyan
    Write-Host '   Victory Road & KeyToKey Watchdog v1.0.3' -ForegroundColor Cyan
    Write-Host '   (Release Version)' -ForegroundColor Cyan
    Write-Host '==========================================' -ForegroundColor Cyan

    # [路徑檢查] 若找不到 KTK 則發出警告
    if (!(Test-Path $KeyToKeyPath)) {
        Write-Host "⚠️ 警告：找不到 KeyToKey 執行檔！" -ForegroundColor Red
        Write-Host "路徑：$KeyToKeyPath" -ForegroundColor Red
        Write-Host "自動重啟功能將失效。" -ForegroundColor Yellow
        Write-Host "請修改腳本中的 `$KeyToKeyPath 變數。" -ForegroundColor Yellow
        Start-Sleep 3
    }

    # Webhook 設定檢查與輸入
    if ([string]::IsNullOrWhiteSpace($DiscordWebhookUrl)) {
        Write-Host ''; Write-Host $Msg_Ask_Webhook -ForegroundColor Yellow
        $InputUrl = Read-Host 'URL'
        if (![string]::IsNullOrWhiteSpace($InputUrl)) {
            $DiscordWebhookUrl = $InputUrl.Trim()
            $DiscordWebhookUrl | Out-File -FilePath $WebhookFile -Encoding UTF8
            Write-Host $Msg_Webhook_Saved -ForegroundColor Green
        }
    }

    # 詢問關機設定 (還原詳細提示語)
    Write-Host ''; Write-Host '[設定] 當遊戲崩潰時，是否要執行電腦關機保護？ (按 Y 啟用，按其他鍵停用)' -ForegroundColor Yellow
    $ShutdownInput = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    $EnableShutdown = ($ShutdownInput.Character -eq 'y' -or $ShutdownInput.Character -eq 'Y')
    if ($EnableShutdown) { Write-Host 'Y (已啟用關機保護)' -ForegroundColor Red } else { Write-Host 'N (僅關閉程式)' -ForegroundColor Green }

    # 詢問按鍵設定
    Write-Host ''; Write-Host '[設定] 請輸入 KTK 啟動熱鍵  [預設: F7]' -ForegroundColor Yellow
    $InputKey = Read-Host '請輸入'
    $TargetKeyName = if ([string]::IsNullOrWhiteSpace($InputKey)) { 'F7' } else { $InputKey.Trim().ToUpper() }
    Write-Host ('已設定按鍵: ' + $TargetKeyName) -ForegroundColor Green

    # 詢問心跳頻率
    Write-Host ''; Write-Host '[設定] 請輸入 Discord 定期回報間隔 (分鐘) [預設: 5 分鐘]' -ForegroundColor Yellow
    $InputInterval = Read-Host '請輸入'
    if (![string]::IsNullOrWhiteSpace($InputInterval) -and ($InputInterval -match '^\d+$')) {
        $Global:HeartbeatInterval = [int]$InputInterval
    }
    Write-Host ('已設定回報間隔: ' + $Global:HeartbeatInterval + ' 分鐘') -ForegroundColor Green

    try { [Console]::CursorVisible = $false } catch {}

    # ==========================================
    # 5. Main Loop（主監控迴圈）
    # ==========================================
    $FreezeCount = 0; $NoResponseCount = 0
    $Global:LastBitmapCache = Capture-ScreenBitmap # 監控開始前先截一張圖作為基準
    $LastPixelData = Get-PixelsFromBitmap $Global:LastBitmapCache

    Write-Host ''; Write-Host '=== 監控開始 (按 Q 停止並回報) ===' -ForegroundColor Cyan
    # 手動清空一次日誌，確保監控開始前的雜訊不會被計入
    $Global:SessionLog = @()
    Send-Discord-Report -Title ($Icon_Start + ' ' + $Msg_Title_Start) -Reason $Msg_Reason_Start -ColorType 'Blue' -IsHeartbeat $true

    # [初始化檢查] 啟動前環境檢查 (不等待倒數)
    Ensure-Game-TopMost
    
    if ((Get-Process -Name 'nie' -ErrorAction SilentlyContinue) -and !(Get-Process -Name 'KeyToKey' -ErrorAction SilentlyContinue)) {
        Write-Log "➤ 初始檢查：KeyToKey 未執行，嘗試啟動..." 'Yellow'
        if (Test-Path $KeyToKeyPath) { 
            Start-Process $KeyToKeyPath
            Write-Log $Msg_Wait_Load 'DarkGray'; Start-Sleep 35
            
            # 抓取視窗並按下熱鍵
            $NewKTK = Get-Process -Name 'KeyToKey' -ErrorAction SilentlyContinue
            if ($NewKTK) {
                [Win32Tools]::ShowWindow($NewKTK.MainWindowHandle, [Win32Tools]::SW_RESTORE) | Out-Null
                [Win32Tools]::SetForegroundWindow($NewKTK.MainWindowHandle) | Out-Null
                Start-Sleep 1; Write-Log ($Msg_Send_Key + ' (' + $TargetKeyName + ')...') 'Cyan'
                Send-Key-Native $TargetKeyName | Out-Null; Start-Sleep 1
            }
            Ensure-Game-TopMost
        }
    }

    while ($true) {
        Ensure-Game-TopMost
        
        # [變數初始化] 先清空狀態字串，避免迴圈剛開始時顯示舊資料
        $StatusStr = ""

        # --- 倒數計時迴圈 ---
        for ($i = $LoopIntervalSeconds; $i -gt 0; $i--) {
            # 1. [每秒執行] Steam 干擾攔截
            # 必須每秒都做，因為 Steam 視窗隨時會跳出來遮擋遊戲
            Suppress-Steam-Window
            
            # [功能] 優雅退場：偵測是否按下 Q 鍵
            if ([Console]::KeyAvailable) {
                $k = [Console]::ReadKey($true)
                if ($k.Key -eq 'Q') {
                    Write-Host "`n"
                    Write-Log $Msg_Title_Stop 'Yellow'
                    $FinalDur = New-TimeSpan -Start $ScriptStartTime -End (Get-Date)
                    $FinalTimeStr = "{0:D2}小時{1:D2}分鐘" -f [int][Math]::Floor($FinalDur.TotalHours), $FinalDur.Minutes
                    Send-Discord-Report -Title ($Icon_Stop + ' ' + $Msg_Title_Stop) -Reason "$Msg_Reason_Stop`n⏱️ **共運行：**$FinalTimeStr" -ColorType 'Grey'
                    Write-Host $Msg_Stop_Monitor -ForegroundColor Green; Read-Host; exit
                }
            }

            # 2. [每 5 秒執行] 檢查遊戲與 KTK 運作狀態
            # [優化原理] 因為 Get-Process 指令比較耗時，如果每秒都查，會導致倒數變慢 (時間膨脹)。改成每 5 秒查一次 (當秒數能被 5 整除時)，其他時間直接顯示舊狀態，既省效能又準時。
            if ($i % 5 -eq 0) {
                $CheckGame = Get-Process -Name 'nie' -ErrorAction SilentlyContinue
                $CheckKTK = Get-Process -Name 'KeyToKey' -ErrorAction SilentlyContinue
                $StatusStr = ''
                if ($CheckGame) { if ($CheckGame.Responding) { $StatusStr += $Msg_Game_Run } else { $StatusStr += $Msg_Game_NoResp + ' ' + $Icon_Warn } } else { $StatusStr += $Msg_Game_Lost + ' ' + $Icon_Cross }
                $StatusStr += ' | '
                if ($CheckKTK)  { $StatusStr += $Msg_KTK_Run } else { $StatusStr += $Msg_KTK_Err + ' ' + $Icon_Warn }
            }

            # 計算進度條百分比 (確保在 $i=1 時進度為 100%)
            if ($LoopIntervalSeconds -gt 1) { $Percent = ($LoopIntervalSeconds - $i) / ($LoopIntervalSeconds - 1) } else { $Percent = 1 }
            $ProgressCount = [int][Math]::Floor($Percent * 20)
            
            # 強制游標起步至少為 1 (讓剛開始倒數時就有動靜，不會一片空白)
            if ($ProgressCount -lt 1) { $ProgressCount = 1 }
            
            # 游標閃爍 (每秒閃爍兩次，每次 0.5 秒)
            # 注意：這裡會顯示 $StatusStr，如果在沒更新的那 4 秒內，它會自動顯示上一次的狀態字串
            for ($blink = 0; $blink -lt 2; $blink++) {
                $BarStr = ''
                # 如果進度條已滿 (20格)，不再閃爍最後一格，確保視覺滿版
                if ($ProgressCount -ge 20) { $BarStr = '=' * 20 } 
                else { if ($ProgressCount -gt 0) { if ($blink -eq 0) { $BarStr = '=' * $ProgressCount } else { $BarStr = '=' * ($ProgressCount - 1) + ' ' } } }
                $Bar = '[' + $BarStr + (' ' * (20 - $BarStr.Length)) + ']'
                # 使用 `r (歸位字元) 覆蓋同一行文字，達成動畫效果
                Write-Host ($CR + $Bar + " 倒數 $i 秒... (按 Q 停止) [ $StatusStr ]       ") -NoNewline -ForegroundColor Gray
                Start-Sleep -Milliseconds 500
            }
        }
        Write-Host '' # 倒數結束換行

        # --- 核心檢測邏輯 (雙重日誌偵測版) ---
        $GameProcess = Get-Process -Name 'nie' -ErrorAction SilentlyContinue
        $KTKProcess = Get-Process -Name 'KeyToKey' -ErrorAction SilentlyContinue
        $ErrorTriggered = $false; $ErrorReason = ''
        
        # 準備變數：擷取當前畫面 (CurrentBitmap)
        $CurrentBitmap = Capture-ScreenBitmap
        $CurrentPixelData = Get-PixelsFromBitmap $CurrentBitmap
        $ReportImages = @() # 準備要傳的圖片路徑陣列

        # ==========================================================================
        # 1. [核心檢測] 程式消失 (Crash)
        #    邏輯：如果抓不到遊戲 Process，代表遊戲已經崩潰或被關閉。
        #          這時立刻去查 Windows 系統日誌，看是不是因為硬體錯誤 (141) 導致的。
        # ==========================================================================
        if (!$GameProcess) { 
            $ErrorTriggered = $true
            
            # [設定範圍] 只搜尋過去 5 分鐘內的錯誤 (確保能抓到剛發生的熱騰騰日誌)
            $TimeLimit = (Get-Date).AddMinutes(-5) 
            
            # ----------------------------------------------------------------------
            # [極速優化] 使用 StartTime 參數 (關鍵!)
            # ----------------------------------------------------------------------
            # 舊寫法：先讀取電腦裡十萬筆歷史紀錄，再用 Where-Object 過濾 -> 耗時 2 分鐘 (卡頓主因)
            # 新寫法：直接告訴系統「我只要 5 分鐘內的」 -> 系統直接給結果 -> 耗時 0.1 秒
            # ----------------------------------------------------------------------
            $SysErrs = Get-WinEvent -FilterHashtable @{LogName='System'; Id=141,4101,117; StartTime=$TimeLimit} -ErrorAction SilentlyContinue
            $AppErrs = Get-WinEvent -FilterHashtable @{LogName='Application'; Id=1001; StartTime=$TimeLimit} -ErrorAction SilentlyContinue | Where-Object { $_.Message -match 'LiveKernelEvent' }
            
            # 將兩邊找到的錯誤合併，並按時間倒序排列 (最新的在最上面)
            $AllErrs = @($SysErrs) + @($AppErrs) | Sort-Object TimeCreated -Descending
            
            if ($AllErrs) {
                # [有找到錯誤]：抓出第一筆最新的錯誤
                $RecentError = $AllErrs | Select-Object -First 1
                
                # [代碼解析] 如果是 1001 (WER)，嘗試分析它是不是偽裝的 141
                if ($RecentError.Id -eq 1001) {
                    if ($RecentError.Message -match '141') { $ErrCode = "LiveKernelEvent (141)" }
                    elseif ($RecentError.Message -match '117') { $ErrCode = "LiveKernelEvent (117)" }
                    elseif ($RecentError.Message -match '1a1') { $ErrCode = "LiveKernelEvent (1a1)" }
                    else { $ErrCode = "LiveKernelEvent (1001)" }
                } else {
                    $ErrCode = $RecentError.Id
                }
                
                # [設定原因] 填寫詳細錯誤原因
                # 這裡修正了排版：移除了冒號後的空格，讓 Log 看起來更整齊
                $SysErrMsg = $Msg_Err_Reason + $ErrCode + ')'
                $ErrorReason = $SysErrMsg
                
                # [第一現場回報] 立刻寫入 Log
                Write-Log ($Icon_Cross + ' 偵測到程式消失，並發現系統錯誤: ' + $ErrCode) 'Red'
            } else { 
                # [沒找到錯誤]：關鍵的 Else！(修復觸發保護空白的問題)
                # 如果系統日誌是乾淨的，代表這是一次普通的閃退 (或是日誌還沒寫入)
                # 我們必須給 $ErrorReason 一個預設值，否則後面的「觸發保護」會顯示空白
                $ErrorReason = $Msg_Err_Crash 
            }
        }

        # 2. 檢測：無回應
        if ($GameProcess -and !$GameProcess.Responding) {
            $NoResponseCount++
            Write-Log ($Icon_Warn + ' ' + $Msg_Warn_NoResp + ' (' + $NoResponseCount + '/' + $NoResponseThreshold + ')') 'Yellow'
            if ($NoResponseCount -ge $NoResponseThreshold) { 
                $ErrorTriggered = $true; $ErrorReason = $Msg_Err_NoResp
                Stop-Process -Name 'nie' -Force -ErrorAction SilentlyContinue 
            }
        } else { $NoResponseCount = 0 }

        # 3. 檢測：凍結
        if ($CurrentPixelData -and $LastPixelData) {
            $Similarity = Get-Similarity $CurrentPixelData $LastPixelData
            if ($Similarity -ge $FreezeSimilarity) {
                $FreezeCount++
                Write-Log ($Icon_Warn + ' ' + $Msg_Warn_Freeze + $Similarity + '%) (' + $FreezeCount + '/' + $FreezeThreshold + ')') 'Yellow'
                
                # 只要偵測到凍結徵兆，就立刻存圖保留證據
                if ($Global:LastBitmapCache) { $PathPrev = Save-BitmapToFile $Global:LastBitmapCache 'Freeze_Prev'; if ($PathPrev) { $ReportImages += $PathPrev } }
                if ($CurrentBitmap) { $PathCurr = Save-BitmapToFile $CurrentBitmap 'Freeze_Curr'; if ($PathCurr) { $ReportImages += $PathCurr } }

                # 判定嚴重程度
                if ($FreezeCount -ge $FreezeThreshold) { 
                    # 達到閾值 (例如 3/3)，判定為死當，觸發紅燈報警與關機
                    $ErrorTriggered = $true; $ErrorReason = $Msg_Err_Freeze 
                } else { 
                    # 未達閾值 (例如 1/3, 2/3)，發送黃色警告訊息 (不附日誌)
                    Send-Discord-Report -Title ($Icon_Warn + ' ' + $Msg_Discord_Title_W) -Reason "畫面相似度過高 ($Similarity%) - 累積 $FreezeCount/$FreezeThreshold" -ColorType 'Yellow' -ImagePaths $ReportImages 
                }
            } else { 
                $FreezeCount = 0 
            }
        }

        # ==========================================================================
        # 4. [輔助檢測] 系統錯誤 (被動查驗模式)
        #    邏輯：當遊戲已經出事了 (前面已經觸發 ErrorTriggered)，我們才來這裡查驗屍體。
        #    目的：解決「有 141 錯誤但遊戲其實沒掛掉」導致的誤殺問題。
        # ==========================================================================
        if ($ErrorTriggered) {
            # 只有當「已經發生異常 ($ErrorTriggered = $true)」時，才執行這段。
            # 如果遊戲活得好好的，就算系統有 141 錯誤 (例如顯卡驅動重置)，我們也假裝沒看到，
            # 讓遊戲繼續跑，絕對不主動殺它。
            
            $TimeLimit = (Get-Date).AddMinutes(-5)
            
            # [極速優化] 同樣使用 StartTime 進行秒讀
            $SysErrs = Get-WinEvent -FilterHashtable @{LogName='System'; Id=141,4101,117; StartTime=$TimeLimit} -ErrorAction SilentlyContinue
            $AppErrs = Get-WinEvent -FilterHashtable @{LogName='Application'; Id=1001; StartTime=$TimeLimit} -ErrorAction SilentlyContinue | Where-Object { $_.Message -match 'LiveKernelEvent' }
            $AllErrs = @($SysErrs) + @($AppErrs) | Sort-Object TimeCreated -Descending
            
            if ($AllErrs) {
                $RecentError = $AllErrs | Select-Object -First 1
                
                # [代碼解析] 
                if ($RecentError.Id -eq 1001) {
                    if ($RecentError.Message -match '141') { $ErrCode = "LiveKernelEvent (141)" }
                    elseif ($RecentError.Message -match '117') { $ErrCode = "LiveKernelEvent (117)" }
                    elseif ($RecentError.Message -match '1a1') { $ErrCode = "LiveKernelEvent (1a1)" }
                    else { $ErrCode = "LiveKernelEvent (1001)" }
                } else {
                    $ErrCode = $RecentError.Id
                }
                
                # 組合錯誤訊息 (移除冒號後空格)
                $SysErrMsg = $Msg_Err_Reason + $ErrCode + ')'
                
                # [靜默補充機制] (解決 "補充偵測" 廢話問題)
                # 檢查現在的 $ErrorReason 裡面，是不是已經包含這個錯誤代碼了？
                # - 如果已經有了：什麼都不做 (避免重複)。
                # - 如果還沒有：偷偷把這行加到 $ErrorReason 變數裡。
                # 注意：這裡不使用 Write-Log！我們只把資訊加進去，留給最後的 Discord 報告一次講完。
                if ($ErrorReason -notmatch [regex]::Escape($ErrCode)) {
                    $ErrorReason += "`n[系統紀錄] $SysErrMsg"
                }
            }
        }
        
        # --- 異常處理流程 (紅色嚴重錯誤) ---
        if ($ErrorTriggered) {
            # 1. 計算時長 (變數供後續使用)
            $FinalDur = New-TimeSpan -Start $ScriptStartTime -End (Get-Date)
            $FinalTimeStr = "{0:D2}小時{1:D2}分鐘" -f [int][Math]::Floor($FinalDur.TotalHours), $FinalDur.Minutes

            # 2. 顯示：❌ 觸發保護
            Write-Log ($Icon_Cross + ' ' + $Msg_Prot_Trig + $ErrorReason) 'Red'
            
            # 3. 顯示：➤ 觸發系統保護 (若有開啟關機)
            if ($EnableShutdown) { Write-Log "➤ 將執行自動關機程序" 'Yellow' }
            
            # 4. 顯示：⏱️ 本次共掛機
            Write-Log "⏱️ 本次共掛機：$FinalTimeStr" 'Cyan'

            # 存取崩潰截圖
            if ($ReportImages.Count -eq 0 -and $CurrentBitmap) {
                $PathCrash = Save-BitmapToFile $CurrentBitmap 'Crash'
                if ($PathCrash) { $ReportImages += $PathCrash }
            }
            # 嘗試關閉 KTK 失敗時捕捉真實錯誤訊息
            if ($KTKProcess) { try { Stop-Process -Name 'KeyToKey' -Force -ErrorAction Stop } catch { Write-Log "⚠️ 無法強制關閉 KeyToKey: $($_.Exception.Message)" 'Yellow' } }
            Stop-Process -Name 'nie' -Force -ErrorAction SilentlyContinue
            
            # 發送 Discord 通知 (加入排版好的關機說明)
            $DiscordReason = if ($EnableShutdown) { "$ErrorReason`n(已執行自動關機程序)" } else { $ErrorReason }
            Send-Discord-Report -Title ($Icon_Cross + ' ' + $Msg_Discord_Title) -Reason $DiscordReason -ColorType 'Red' -ImagePaths $ReportImages
            
            # 資源釋放
            if ($Global:LastBitmapCache) { $Global:LastBitmapCache.Dispose() }
            if ($CurrentBitmap) { $CurrentBitmap.Dispose() }

            if ($EnableShutdown) { 
                $GuiResult = Show-Crash-Warning-GUI -Reason $ErrorReason
                if ($GuiResult -eq [System.Windows.Forms.DialogResult]::OK) {
                    Write-Log $Msg_Shutdown 'Red'; Stop-Computer -Force; exit 
                } else {
                    Write-Log "使用者已取消關機。" 'Yellow'
                    Write-Log $Msg_Stop_Monitor 'Red'; Read-Host; exit
                }
            } else {
                Write-Log $Msg_Stop_Monitor 'Red'; Read-Host; exit
            }
        }

        # 釋放舊的緩存，把 Current 變成 Last，供下次迴圈比對
        if ($Global:LastBitmapCache) { $Global:LastBitmapCache.Dispose() }
        $Global:LastBitmapCache = $CurrentBitmap # 轉移物件所有權，不要 Dispose Current
        $LastPixelData = $CurrentPixelData
        
        # 狀態顯示
        if (!$ErrorTriggered -and $KTKProcess) { Write-Log ('➤ ' + $Msg_Status_OK) 'DarkGray' }

        # --- 定期心跳 (Heartbeat) ---
        $TimeSinceLastHeartbeat = (Get-Date) - $Global:LastHeartbeatTime
        if ($TimeSinceLastHeartbeat.TotalMinutes -ge $Global:HeartbeatInterval) {
            # 1. 準備時間字串
            $HbDur = New-TimeSpan -Start $ScriptStartTime -End (Get-Date)
            $HbTimeStr = "{0:D2}小時{1:D2}分鐘" -f [int][Math]::Floor($HbDur.TotalHours), $HbDur.Minutes
            
            # 2. 先寫入日誌 (Console + Memory)
            Write-Log ('➤ ' + $Msg_Sent_Report + " (已運行時間：$HbTimeStr)") 'Cyan'
            
            # 3. 再發送 Discord
            $HbPath = Save-BitmapToFile $Global:LastBitmapCache 'Heartbeat'
            $HbPaths = if ($HbPath) { @($HbPath) } else { @() }
            Send-Discord-Report -Title ($Icon_Heart + ' ' + $Msg_Discord_HB) -Reason 'Heartbeat' -ColorType 'Green' -ImagePaths $HbPaths -IsHeartbeat $true
            $Global:LastHeartbeatTime = Get-Date
        }

        # --- KTK 自動修復 (如果遊戲還在但腳本掛了) ---
        if ($GameProcess -and !$KTKProcess) {
            Write-Log ('➤ ' + $Msg_KTK_Restart) 'White' $true
            if (Test-Path $KeyToKeyPath) {
                Start-Process $KeyToKeyPath; Write-Log $Msg_Wait_Load 'DarkGray'; Start-Sleep 35
                $NewKTK = Get-Process -Name 'KeyToKey' -ErrorAction SilentlyContinue
                if ($NewKTK) {
                    [Win32Tools]::ShowWindow($NewKTK.MainWindowHandle, [Win32Tools]::SW_RESTORE) | Out-Null
                    [Win32Tools]::SetForegroundWindow($NewKTK.MainWindowHandle) | Out-Null
                    Start-Sleep 1; Write-Log ($Msg_Send_Key + ' (' + $TargetKeyName + ')...') 'Cyan'
                    Send-Key-Native $TargetKeyName | Out-Null; Start-Sleep 1
                }
                Ensure-Game-TopMost; Write-Log ($Icon_Check + ' ' + $Msg_Recovered) 'Green'
            }
        }
    }
} catch {
    Write-Host "`n[嚴重錯誤] 程式發生未預期的例外狀況：" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host "按 Enter 鍵離開..." -NoNewline; Read-Host
}