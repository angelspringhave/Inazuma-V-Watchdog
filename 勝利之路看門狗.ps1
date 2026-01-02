<#
    【勝利之路 & KeyToKey 看門狗 v1.6.4】
    Fix:修正並美化了關機警告彈出視窗
    Fix:因為檢測的邏輯順序而發生如果是因為GPU問題導致遊戲崩潰卻不會顯示GPU有問題的歷史紀錄
    Update:在關閉 Steam 視窗之後，立刻強制把焦點抓回遊戲身上
    Add:關機程序紀錄：執行關機前，會於控制台 Log 與 Discord 通知中明確標註執行自動關機
    Update:當觸發關機嘗試關閉 KTK時，若因權限問題失敗則跳過，不影響後續關機流程
    Fix:修復崩潰時圖片檔案裡截圖沒有自動刪除的問題
#>

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

    $Icon_Warn   = [char]0x26A0 # ⚠️
    $Icon_Check  = [char]0x2705 # ✅
    $Icon_Cross  = [char]0x274C # ❌
    $Icon_Heart  = [string][char]0xD83E + [char]0xDEC0 # 🫀
    $Icon_Bullet = [char]0x2022 # •
    $Icon_Start  = [string][char]0x26A1 + [char]0x26BD # ⚡⚽

    # --- 中文訊息設定 (集中管理，方便修改) ---
    $Msg_Title_Start    = '看門狗 v1.6.4 已啟動'
    $Msg_Reason_Start   = '啟動通知'
    $Msg_Game_Run       = '勝利之路：運作中'
    $Msg_Game_NoResp    = '勝利之路：無回應'
    $Msg_Game_Lost      = '勝利之路：消失'
    $Msg_KTK_Run        = 'KeyToKey：運作中'
    $Msg_KTK_Err        = 'KeyToKey：異常'
    $Msg_Warn_NoResp    = '警告：程式無回應'
    $Msg_Warn_Freeze    = '警告：畫面凍結 (相似度 '
    $Msg_Err_Freeze     = '畫面凍結 (死機)'
    $Msg_Err_NoResp     = '程式無回應 (卡死)'
    $Msg_Err_Crash      = '程式崩潰 (消失)'
    $Msg_Err_Sys        = '偵測到系統錯誤： ID'
    $Msg_Err_Reason     = '系統嚴重錯誤 (ID:'
    $Msg_Prot_Trig      = '觸發保護：'
    $Msg_Discord_Title  = '嚴重異常終止'
    $Msg_Discord_Title_W= '異常徵兆警告'
    $Msg_Discord_HB     = '看門狗定期報告'
    $Msg_Discord_HBTxt  = '定期健康報告'
    $Msg_Discord_SysOK  = '勝利之路：🟢    |   KeyToKey：🟢 '
    $Msg_Discord_Log    = '近期紀錄：'
    $Msg_Shutdown       = '60秒後關機...'
    $Msg_GUI_Title      = '⚠️ 掛機失敗——關機預警'
    $Msg_GUI_Cancel     = '🚫 取消關機'
    $Msg_Stop_Monitor   = '監控已停止。請按 Enter 鍵離開視窗...'
    $Msg_Status_OK      = '掛機運作中'
    $Msg_Sent_Report    = '已發送定期 Discord 報告'
    $Msg_KTK_Restart    = 'KeyToKey 重啟中...'
    $Msg_Wait_Load      = '等待 35 秒載入...'
    $Msg_Send_Key       = '發送按鍵'
    $Msg_Recovered      = '復原完畢'
    $Msg_Footer_Base    = 'Watchdog v1.6.4'
    $Msg_Ask_Webhook    = '[設定] 初次執行，請輸入 Discord Webhook 網址 (輸入完畢按 Enter):'
    $Msg_Webhook_Saved  = '網址已儲存至 webhook.txt，下次將自動讀取。'

    $ScriptStartTime = Get-Date
    $DiscordUserID   = '649980145020436497' 

    # ==========================================
    # 1. User Settings（使用者設定區）
    # ==========================================
    $KeyToKeyPath = 'D:\Users\user\Downloads\KeyToKey\KeyToKey.exe'
    $ScreenshotDir = "D:\Users\user\Desktop\勝利之路看門狗"
    $LogSavePath = $env:USERPROFILE + '\Desktop\Watchdog_Log_Latest.txt'

    # Webhook 讀取邏輯
    $ScriptPath = $MyInvocation.MyCommand.Path
    $ScriptDir  = Split-Path $ScriptPath -Parent
    $WebhookFile = Join-Path $ScriptDir 'webhook.txt'
    if (Test-Path $WebhookFile) {
        $DiscordWebhookUrl = (Get-Content $WebhookFile -Raw).Trim()
    } else {
        $DiscordWebhookUrl = ''
    }

    # 監控參數
    $LoopIntervalSeconds = 75  # 每次檢測的間隔秒數 # 降為 75 秒，加上處理時間後，總間隔會接近 90 秒 (1:30)
    $FreezeThreshold = 3       # 連續畫面凍結幾次才判定為當機
    $NoResponseThreshold = 3   # 連續無回應幾次才判定為卡死
    $FreezeSimilarity = 98.5   # 畫面相似度閾值 (%)

    # 初始化全域變數
    $Global:SessionLog = @()
    $Global:LastReportLogIndex = 0 # 紀錄上次回報到 Log 的哪一行
    $Global:LastHeartbeatTime = Get-Date
    $Global:HeartbeatInterval = 10 # 分鐘
    # 初始化上一張畫面的緩存 (用於凍結對比)
    $Global:LastBitmapCache = $null 

    # ==========================================
    # 2. System Core（系統核心與 Windows API）
    # ==========================================
    # 設定 Console 緩衝區大小，避免文字太長被截斷
    try {
        $PSWindow = (Get-Host).UI.RawUI
        $BufferSize = $PSWindow.BufferSize; $BufferSize.Width = 120; $PSWindow.BufferSize = $BufferSize
        $WindowSize = $PSWindow.WindowSize; $WindowSize.Width = 120; $PSWindow.WindowSize = $WindowSize
    } catch {}

    if (!(Test-Path $ScreenshotDir)) { New-Item -ItemType Directory -Path $ScreenshotDir | Out-Null }
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
        [DllImport("user32.dll")] public static extern IntPtr LoadKeyboardLayout(string pwszKLID, uint Flags);
        [DllImport("user32.dll")] public static extern bool ActivateKeyboardLayout(IntPtr hkl, uint Flags);
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
                // 只檢查可見的視窗
                if (IsWindowVisible(hWnd)) {
                    int pid;
                    GetWindowThreadProcessId(hWnd, out pid);
                    try {
                        Process p = Process.GetProcessById(pid);
                        string pName = p.ProcessName.ToLower();

                        // 檢查處理程序名稱是否為 steam 或 steamwebhelper (負責顯示網頁內容的程式)
                        if (pName == "steam" || pName == "steamwebhelper") {
                            StringBuilder sb = new StringBuilder(256);
                            GetWindowText(hWnd, sb, 256);
                            string title = sb.ToString();

                            // 排除沒有標題的隱藏視窗 (系統繪圖用)，只抓真正跳出來的視窗
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

        // 發送關閉訊號
        public static void CloseWindow(IntPtr hWnd) {
            PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
        }
    }
'@
    Add-Type -TypeDefinition $Win32Code

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
        
        # [v1.6.1] 修正換行邏輯，確保倒數計時被中斷時能正確換行顯示
        if ($ForceNewLine) { 
            Write-Host '' 
        } 
        Write-Host ($LogLine + '          ') -ForegroundColor $Color
        $Global:SessionLog += $LogLine
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
    # 函式：Ensure-English-IME 
    # 功能：強制切換輸入法為英文 (0x0409)，避免 KTK 送出注音/拼音
    # -----------------------------------------------------------
    function Ensure-English-IME {
        try {
            # 00000409 代表美式英文，1 (KLF_ACTIVATE) 代表立即啟用
            $HKL = [Win32Tools]::LoadKeyboardLayout("00000409", 1) 
            [Win32Tools]::ActivateKeyboardLayout($HKL, 0) | Out-Null
        } catch {}
    }

    # -----------------------------------------------------------
    # 函式：Send-Key-Native
    # 功能：使用 SendKeys 模擬鍵盤按鍵輸入 (用於重啟 KTK)
    # -----------------------------------------------------------
    function Send-Key-Native ($KeyName) {
        try {
            $KeyStr = '{' + $KeyName + '}'
            [System.Windows.Forms.SendKeys]::SendWait($KeyStr)
            return $true
        } catch { return $false }
    }

    # -----------------------------------------------------------
    # 函式：Show-Crash-Warning-GUI (v1.6.3 最終美化版)
    # 功能：顯示現代化暗黑風格的警告視窗，倒數 60 秒後自動關機
    # -----------------------------------------------------------
    function Show-Crash-Warning-GUI {
        param([string]$Reason)
        
        # --- 1. 定義符號 (使用 Unicode 代碼，解決亂碼問題) ---
        $Sym_Warn   = [char]0x26A0  # ⚠️
        $Sym_Cancel = [char]0x2716  # ✖
        
        # --- 2. 定義配色 (暗黑警報主題) ---
        $Color_Bg        = [System.Drawing.Color]::FromArgb(30, 30, 30)   # 深灰背景
        $Color_Accent    = [System.Drawing.Color]::FromArgb(255, 60, 60)  # 亮警報紅
        $Color_TextPri   = [System.Drawing.Color]::White              # 純白文字
        $Color_TextSec   = [System.Drawing.Color]::FromArgb(200, 200, 200)# 淺灰文字
        $Color_BtnBg     = [System.Drawing.Color]::White              # 按鈕背景(白)
        $Color_BtnText   = [System.Drawing.Color]::Black              # 按鈕文字(黑)

        # --- 3. 表單設定 ---
        $FormW = 600
        $FormH = 380
        
        $Form = New-Object System.Windows.Forms.Form
        $Form.Size = New-Object System.Drawing.Size($FormW, $FormH)
        $Form.StartPosition = 'CenterScreen'
        $Form.TopMost = $true
        $Form.FormBorderStyle = 'None' # 無邊框
        $Form.BackColor = $Color_Accent 
        $Form.Padding = New-Object System.Windows.Forms.Padding(4) # 紅色邊框

        $MainPanel = New-Object System.Windows.Forms.Panel
        $MainPanel.Dock = 'Fill'
        $MainPanel.BackColor = $Color_Bg
        $Form.Controls.Add($MainPanel)

        # --- 4. UI 元件 (使用自動置中) ---
        
        # [標題]
        $LblTitle = New-Object System.Windows.Forms.Label
        $LblTitle.Text = "$Sym_Warn 偵測到嚴重錯誤"
        $LblTitle.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 20, [System.Drawing.FontStyle]::Bold)
        $LblTitle.ForeColor = $Color_Accent
        $LblTitle.AutoSize = $false
        $LblTitle.Size = New-Object System.Drawing.Size([int]($FormW - 8), 50)
        $LblTitle.Location = New-Object System.Drawing.Point(0, 30)
        $LblTitle.TextAlign = 'MiddleCenter' 
        $MainPanel.Controls.Add($LblTitle)

        # [原因]
        $LblReason = New-Object System.Windows.Forms.Label
        $DispReason = if ($Reason.Length -gt 45) { $Reason.Substring(0, 42) + "..." } else { $Reason }
        $LblReason.Text = "$DispReason"
        $LblReason.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 12)
        $LblReason.ForeColor = $Color_TextSec
        $LblReason.AutoSize = $false
        $LblReason.Size = New-Object System.Drawing.Size([int]($FormW - 8), 30)
        $LblReason.Location = New-Object System.Drawing.Point(0, 80)
        $LblReason.TextAlign = 'MiddleCenter'
        $MainPanel.Controls.Add($LblReason)

        # [倒數數字]
        $LblCount = New-Object System.Windows.Forms.Label
        $LblCount.Text = "60"
        $LblCount.Font = New-Object System.Drawing.Font("Arial", 55, [System.Drawing.FontStyle]::Bold)
        $LblCount.ForeColor = $Color_TextPri
        $LblCount.AutoSize = $false
        $LblCount.Size = New-Object System.Drawing.Size([int]($FormW - 8), 100)
        $LblCount.Location = New-Object System.Drawing.Point(0, 115)
        $LblCount.TextAlign = 'MiddleCenter'
        $MainPanel.Controls.Add($LblCount)
        
        # [倒數文字]
        $LblSub = New-Object System.Windows.Forms.Label
        $LblSub.Text = "秒後將執行系統保護關機..."
        $LblSub.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 11)
        $LblSub.ForeColor = $Color_Accent
        $LblSub.AutoSize = $false
        $LblSub.Size = New-Object System.Drawing.Size([int]($FormW - 8), 30)
        $LblSub.Location = New-Object System.Drawing.Point(0, 215)
        $LblSub.TextAlign = 'TopCenter'
        $MainPanel.Controls.Add($LblSub)

        # [取消按鈕]
        $BtnCancel = New-Object System.Windows.Forms.Button
        $BtnCancel.Text = "$Sym_Cancel 取消關機"
        $BtnCancel.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 16, [System.Drawing.FontStyle]::Bold)
        $BtnCancel.Size = New-Object System.Drawing.Size(260, 60)
        
        # 重新計算按鈕置中 (強制轉型 int 避免崩潰)
        $BtnX = [int](($FormW - 260) / 2)
        $BtnCancel.Location = New-Object System.Drawing.Point([int]($BtnX - 4), 270)
        
        $BtnCancel.BackColor = $Color_BtnBg
        $BtnCancel.ForeColor = $Color_BtnText
        $BtnCancel.FlatStyle = 'Flat'
        $BtnCancel.Cursor = [System.Windows.Forms.Cursors]::Hand
        $BtnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $MainPanel.Controls.Add($BtnCancel)

        # --- 5. Timer 邏輯 ---
        $Timer = New-Object System.Windows.Forms.Timer
        $Timer.Interval = 1000
        $Script:CountDown = 60
        $Timer.Add_Tick({
            $Script:CountDown--
            $LblCount.Text = "$Script:CountDown"
            if ($Script:CountDown -le 0) {
                $Timer.Stop()
                $Form.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $Form.Close()
            }
        })
        $Timer.Start()

        # 預設聚焦取消按鈕，方便直接按 Enter/空白鍵取消
        $Form.Add_Shown({ $BtnCancel.Focus() })

        $Result = $Form.ShowDialog()
        $Timer.Stop()
        $Form.Dispose()
        return $Result
    }

    # -----------------------------------------------------------
    # 函式：Send-Discord-Report
    # 功能：組合訊息、日誌與截圖，發送 Embed 到 Discord
    # 參數：
    #   $ColorType - 決定 Embed 邊條顏色 (Green, Red, Yellow, Blue)
    #   $ImagePaths - 圖片路徑陣列 (支援多張圖)
    #   $IsHeartbeat - 是否為定期報告 (影響日誌抓取範圍)
    # -----------------------------------------------------------
    function Send-Discord-Report {
        param(
            [string]$Title, 
            [string]$Reason, 
            [string]$ColorType='Green', 
            [string[]]$ImagePaths=@(), 
            [bool]$IsHeartbeat=$false
        )
        if ([string]::IsNullOrWhiteSpace($DiscordWebhookUrl) -or $DiscordWebhookUrl -eq 'YOUR_WEBHOOK_HERE') { return }
        
        # 如果不是心跳包 (例如黃色警告)，不要強制換行，避免空行問題
        # 原本是 $true (強制換行)，現在改為 $false，因為前面已經有一行警告訊息了
        if (!$IsHeartbeat) { Write-Log 'Uploading Report...' 'Cyan' $false }

        # --- 日誌處理邏輯 ---
        $LogPreviewLines = @()
        
        if ($IsHeartbeat) {
            # 定期報告：只抓取新紀錄
            $NewCount = $Global:SessionLog.Count
            if ($NewCount -gt $Global:LastReportLogIndex) {
                for ($k = $Global:LastReportLogIndex; $k -lt $NewCount; $k++) {
                    $LogPreviewLines += $Global:SessionLog[$k]
                }
            }
            $Global:LastReportLogIndex = $NewCount
        } else {
            # 異常報告：抓取最後 15 行
            $MaxLines = 15
            $Count = 0
            for ($k = $Global:SessionLog.Count - 1; $k -ge 0; $k--) {
                $LogPreviewLines += $Global:SessionLog[$k]
                $Count++
                if ($Count -ge $MaxLines) { break }
            }
            [array]::Reverse($LogPreviewLines)
        }

        $LogPreview = $LogPreviewLines -join $LF
        if ([string]::IsNullOrWhiteSpace($LogPreview)) { $LogPreview = '(無)' }

        $Global:SessionLog | Out-File -FilePath $LogSavePath -Encoding UTF8

        # --- 建立 Embed ---
        $ColorMap = @{ 'Green'=5763719; 'Red'=15548997; 'Yellow'=16705372; 'Blue'=5793266 }
        
        $Duration = New-TimeSpan -Start $ScriptStartTime -End (Get-Date)
        $RunTimeStr = "{0:D2}小時{1:D2}分鐘" -f [int][Math]::Floor($Duration.TotalHours), $Duration.Minutes

        $DescHeader = ''
        $MentionContent = ''

        if ($IsHeartbeat) {
            $DescHeader = '**' + $Icon_Check + ' ' + $Msg_Discord_HBTxt + '**' + $LF + 
                          $Msg_Discord_SysOK + $LF + 
                          "(每 $Global:HeartbeatInterval 分鐘回報一次)" + $LF + $LF + 
                          '⏱️ **已運行時間**' + $LF + 
                          $RunTimeStr
        } else {
            # [v1.6.1] 修正排版：將「原因」與「掛機時長」分段顯示
            $DescHeader = "**異常原因：**$LF" + $Reason + $LF + $LF + 
                          "⏳ **已掛機：**$LF" + $RunTimeStr
            
            # [v1.6] 紅燈與黃燈都要 @使用者
            if ($ColorType -eq 'Red' -or $ColorType -eq 'Yellow') { $MentionContent = "<@$DiscordUserID>" }
        }

        # 如果是黃色警告，只顯示原因，不顯示日誌預覽區塊 (避免洗版)
        if ($ColorType -ne 'Yellow') {
            $EmbedDesc = $DescHeader + $LF + $LF + '**📋 ' + $Msg_Discord_Log + '**' + $LF + '```' + $LF + $LogPreview + $LF + '```'
        } else {
            $EmbedDesc = $DescHeader 
        }

        $FooterTxt = $Msg_Footer_Base + ' ' + $Icon_Bullet + ' ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

        $Embed = @{
            title = $Title
            description = $EmbedDesc
            color = $ColorMap[$ColorType]
            footer = @{ text = $FooterTxt }
        }
        $Payload = @{ content = $MentionContent; embeds = @($Embed) }
        
        # 轉為 JSON 並確保 UTF-8 編碼
        $JsonPayload = $Payload | ConvertTo-Json -Depth 10 -Compress

        # --- 發送 Multipart HTTP 請求 ---
        $HttpClient = New-Object System.Net.Http.HttpClient
        $Streams = @() 
        
        try {
            $Form = New-Object System.Net.Http.MultipartFormDataContent
            $Enc = [System.Text.Encoding]::UTF8
            
            $Form.Add((New-Object System.Net.Http.StringContent($JsonPayload, $Enc, 'application/json')), 'payload_json')

            $ImgIndex = 1
            foreach ($Path in $ImagePaths) {
                if (![string]::IsNullOrEmpty($Path) -and (Test-Path $Path)) {
                    $FS = [System.IO.File]::OpenRead($Path)
                    $Streams += $FS
                    $ImgContent = New-Object System.Net.Http.StreamContent($FS)
                    $ImgContent.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse('image/png')
                    $Form.Add($ImgContent, "file$ImgIndex", [System.IO.Path]::GetFileName($Path))
                    $ImgIndex++
                }
            }

            # 僅在「非心跳」且「非黃色警告」時，才附上完整 Log 文字檔
            if (!$IsHeartbeat -and $ColorType -ne 'Yellow' -and (Test-Path $LogSavePath)) {
                $FS2 = [System.IO.File]::OpenRead($LogSavePath)
                $Streams += $FS2
                $TxtContent = New-Object System.Net.Http.StreamContent($FS2)
                $TxtContent.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse('text/plain')
                $Form.Add($TxtContent, "file_log", 'Watchdog_Log.txt')
            }

            # [v1.6] 加入 charset=utf-8 確保中文顯示正常
            $HttpClient.PostAsync($DiscordWebhookUrl, $Form).Result | Out-Null
            
        } catch {
            Write-Log "Discord 上傳失敗: $_" 'Red' $true
        } finally {
            foreach ($s in $Streams) { $s.Close(); $s.Dispose() }
            if ($HttpClient) { $HttpClient.Dispose() }
            if ($Form) { $Form.Dispose() }
        }

        # 修復：等待 1 秒確保 HttpClient 徹底釋放檔案，再執行刪除
        Start-Sleep -Seconds 1
        
        # 刪除暫存截圖
        foreach ($Path in $ImagePaths) {
            if (Test-Path $Path) { try { Remove-Item $Path -Force -ErrorAction SilentlyContinue } catch {} }
        }
        # 確保日誌也被刪除
        if (Test-Path $LogSavePath) {
            try { Remove-Item $LogSavePath -Force -ErrorAction SilentlyContinue } catch {}
        }
    }

    # -----------------------------------------------------------
    # 函式：Suppress-Steam-Window (v1.6.4 焦點修正版)
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

        # [修正] 如果有關閉過視窗，立刻把焦點搶回遊戲，避免 KTK 按錯地方
        if ($ClosedAny) {
            Ensure-Game-TopMost
        }
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
            for ($x=0; $x -lt 12; $x++) {
                for ($y=0; $y -lt 12; $y++) {
                    $Pixels[$x, $y] = $Small.GetPixel($x, $y).ToArgb()
                }
            }
            $Small.Dispose()
            return ,$Pixels 
        } catch { return $null }
    }

    # -----------------------------------------------------------
    # 函式：Get-Similarity
    # 功能：比對兩組像素矩陣的相似度 (0~100%)
    # -----------------------------------------------------------
    function Get-Similarity ($PixA, $PixB) {
        if (!$PixA -or !$PixB) { return 0 }
        $Match = 0; $Total = 144
        for ($x=0; $x -lt 12; $x++) {
            for ($y=0; $y -lt 12; $y++) {
                $valA = $PixA[$x, $y]; $valB = $PixB[$x, $y]
                $R1 = ($valA -shr 16) -band 255; $G1 = ($valA -shr 8) -band 255; $B1 = $valA -band 255
                $R2 = ($valB -shr 16) -band 255; $G2 = ($valB -shr 8) -band 255; $B2 = $valB -band 255
                
                if ([Math]::Abs($R1 - $R2) -lt 20 -and [Math]::Abs($G1 - $G2) -lt 20 -and [Math]::Abs($B1 - $B2) -lt 20) { 
                    $Match++ 
                }
            }
        }
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
    Write-Host '   Victory Road & KeyToKey Watchdog v1.6.4' -ForegroundColor Cyan
    Write-Host '   (Dual-Image Freeze Detect)' -ForegroundColor Cyan
    Write-Host '==========================================' -ForegroundColor Cyan

    # Webhook 設定檢查與輸入
    if ([string]::IsNullOrWhiteSpace($DiscordWebhookUrl)) {
        Write-Host ''
        Write-Host $Msg_Ask_Webhook -ForegroundColor Yellow
        $InputUrl = Read-Host 'URL'
        if (![string]::IsNullOrWhiteSpace($InputUrl)) {
            $DiscordWebhookUrl = $InputUrl.Trim()
            $DiscordWebhookUrl | Out-File -FilePath $WebhookFile -Encoding UTF8
            Write-Host $Msg_Webhook_Saved -ForegroundColor Green
        }
    }

    # 詢問關機設定
    Write-Host ''
    Write-Host '[設定] 當遊戲崩潰時，是否要執行電腦關機保護？ (按 Y 啟用，按其他鍵停用)' -ForegroundColor Yellow
    $ShutdownInput = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    $EnableShutdown = ($ShutdownInput.Character -eq 'y' -or $ShutdownInput.Character -eq 'Y')
    if ($EnableShutdown) { Write-Host 'Y (已啟用關機保護)' -ForegroundColor Red } else { Write-Host 'N (僅關閉程式)' -ForegroundColor Green }

    # 詢問按鍵設定
    Write-Host ''
    Write-Host '[設定] 請輸入 KTK 啟動熱鍵  [直接按 Enter 預設: F7]' -ForegroundColor Yellow
    $InputKey = Read-Host '請輸入'
    $TargetKeyName = if ([string]::IsNullOrWhiteSpace($InputKey)) { 'F7' } else { $InputKey.Trim().ToUpper() }
    Write-Host ('已設定按鍵: ' + $TargetKeyName) -ForegroundColor Green

    # 詢問心跳頻率
    Write-Host ''
    Write-Host '[設定] 請輸入 Discord 定期回報間隔 (分鐘) [按 Enter 預設: 10]' -ForegroundColor Yellow
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

    Write-Host ''
    Write-Host '=== 監控開始 (按 Ctrl+C 停止) ===' -ForegroundColor Cyan
    
    # 手動清空一次日誌，確保監控開始前的雜訊不會被計入
    $Global:SessionLog = @()
    Send-Discord-Report -Title ($Icon_Start + ' ' + $Msg_Title_Start) -Reason $Msg_Reason_Start -ColorType 'Blue' -IsHeartbeat $true

    while ($true) {
        Ensure-Game-TopMost
        Ensure-English-IME
        
        # --- 倒數計時動畫 ---
        for ($i = $LoopIntervalSeconds; $i -gt 0; $i--) {
            
            # [v1.6] 將 Steam 攔截移入倒數迴圈，實現每秒監控
            Suppress-Steam-Window

            # 計算進度百分比 (確保在 $i=1 時進度為 100%)
            if ($LoopIntervalSeconds -gt 1) {
                $Percent = ($LoopIntervalSeconds - $i) / ($LoopIntervalSeconds - 1)
            } else { $Percent = 1 }
            $ProgressCount = [int][Math]::Floor($Percent * 20)
            
            # 檢查程式存活狀態以顯示在進度條上
            $CheckGame = Get-Process -Name 'nie' -ErrorAction SilentlyContinue
            $CheckKTK = Get-Process -Name 'KeyToKey' -ErrorAction SilentlyContinue
            $StatusStr = ''
            if ($CheckGame) { 
                if ($CheckGame.Responding) { $StatusStr += $Msg_Game_Run } else { $StatusStr += $Msg_Game_NoResp + ' ' + $Icon_Warn }
            } else { $StatusStr += $Msg_Game_Lost + ' ' + $Icon_Cross }
            $StatusStr += ' | '
            if ($CheckKTK)  { $StatusStr += $Msg_KTK_Run } else { $StatusStr += $Msg_KTK_Err + ' ' + $Icon_Warn }

            # 游標閃爍效果 (每秒閃爍兩次)
            for ($blink = 0; $blink -lt 2; $blink++) {
                $BarStr = ''
                # 如果進度條已滿 (20格)，不再閃爍最後一格，確保視覺滿版
                if ($ProgressCount -ge 20) {
                    $BarStr = '=' * 20
                } else {
                    if ($ProgressCount -gt 0) { 
                        if ($blink -eq 0) { $BarStr = '=' * $ProgressCount } 
                        else { $BarStr = '=' * ($ProgressCount - 1) + ' ' } 
                    }
                }
                
                $Bar = '[' + $BarStr + (' ' * (20 - $BarStr.Length)) + ']'
                # 使用 `r (歸位) 覆蓋同一行文字
                Write-Host ($CR + $Bar + " 倒數 $i 秒... [ $StatusStr ]       ") -NoNewline -ForegroundColor Gray
                Start-Sleep -Milliseconds 500
            }
        }
        Write-Host '' # 倒數結束換行

        # --- 核心檢測邏輯 ---
        $GameProcess = Get-Process -Name 'nie' -ErrorAction SilentlyContinue
        $KTKProcess = Get-Process -Name 'KeyToKey' -ErrorAction SilentlyContinue
        $ErrorTriggered = $false; $ErrorReason = ''
        
        # 準備變數：擷取當前畫面 (CurrentBitmap)
        $CurrentBitmap = Capture-ScreenBitmap
        $CurrentPixelData = Get-PixelsFromBitmap $CurrentBitmap
        $ReportImages = @() # 準備要傳的圖片路徑陣列

        # 1. 檢測：程式是否消失 (Process 遺失)
        if (!$GameProcess) { 
            $ErrorTriggered = $true
            
            # [v1.6.4 優化] 發現崩潰時，優先進行「驗屍」：檢查剛剛是否有系統硬體錯誤 (如 141)
            # 往前多追溯 30 秒，確保能抓到導致崩潰的瞬間
            $TimeLimit = (Get-Date).AddSeconds(-($LoopIntervalSeconds + 30))
            $KernelErrors = Get-WinEvent -FilterHashtable @{LogName='System'; Id=141,4101,41,117} -ErrorAction SilentlyContinue | Where-Object { $_.TimeCreated -gt $TimeLimit }
            
            if ($KernelErrors) {
                # 如果找到系統錯誤，優先使用系統錯誤作為原因
                $RecentError = $KernelErrors | Select-Object -First 1
                $ErrorReason = $Msg_Err_Reason + ' ' + $RecentError.Id + ')'
                # 在 Log 裡也特別標註一下
                Write-Log ($Icon_Cross + ' 偵測到程式消失，且發現系統錯誤 ID: ' + $RecentError.Id) 'Red'
            } else {
                # 沒找到系統錯誤，才判定為一般崩潰
                $ErrorReason = $Msg_Err_Crash 
            }
        }

        # 2. 檢測：程式是否無回應 (Not Responding)
        if ($GameProcess -and !$GameProcess.Responding) {
            $NoResponseCount++
            Write-Log ($Icon_Warn + ' ' + $Msg_Warn_NoResp + ' (' + $NoResponseCount + '/' + $NoResponseThreshold + ')') 'Yellow'
            if ($NoResponseCount -ge $NoResponseThreshold) { 
                $ErrorTriggered = $true; $ErrorReason = $Msg_Err_NoResp
                Stop-Process -Name 'nie' -Force -ErrorAction SilentlyContinue 
            }
        } else { $NoResponseCount = 0 }

        # 3. 檢測：畫面是否凍結 (像素比對)
        if ($CurrentPixelData -and $LastPixelData) {
            $Similarity = Get-Similarity $CurrentPixelData $LastPixelData
            if ($Similarity -ge $FreezeSimilarity) {
                $FreezeCount++
                Write-Log ($Icon_Warn + ' ' + $Msg_Warn_Freeze + $Similarity + '%) (' + $FreezeCount + '/' + $FreezeThreshold + ')') 'Yellow'
                
                # 只要偵測到凍結徵兆，就立刻存圖保留證據
                if ($Global:LastBitmapCache) {
                    $PathPrev = Save-BitmapToFile $Global:LastBitmapCache 'Freeze_Prev'
                    if ($PathPrev) { $ReportImages += $PathPrev }
                }
                if ($CurrentBitmap) {
                    $PathCurr = Save-BitmapToFile $CurrentBitmap 'Freeze_Curr'
                    if ($PathCurr) { $ReportImages += $PathCurr }
                }

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

        # 4. 檢測：Windows 系統錯誤事件 (Event Log)
        # 即使前面已經觸發(例如凍結)，也要檢查是否有系統錯誤，因為那通常是根本原因
        $TimeLimit = (Get-Date).AddSeconds(-($LoopIntervalSeconds + 30)) # 擴大搜尋範圍
        # 增加偵測 ID: 10016 (DCOM權限, 常見當機前兆)
        $KernelErrors = Get-WinEvent -FilterHashtable @{LogName='System'; Id=141,4101,41,117,10016} -ErrorAction SilentlyContinue | Where-Object { $_.TimeCreated -gt $TimeLimit }
        
        if ($KernelErrors) {
            $RecentError = $KernelErrors | Select-Object -First 1
            $SysErrMsg = $Msg_Err_Reason + ' ' + $RecentError.Id + ')'
            
            if ($ErrorTriggered) {
                # 如果已經有錯誤(例如畫面凍結)，這很有可能是主因，將其追加到原因中
                $ErrorReason += "`n[系統紀錄] $SysErrMsg"
                # 在 Log 中補上一筆紅色紀錄
                Write-Log ($Icon_Cross + ' 補充偵測：' + $SysErrMsg) 'Red'
            } else {
                # 如果還沒觸發，這就是主因
                $ErrorTriggered = $true
                $ErrorReason = $SysErrMsg
                Write-Log ($Icon_Cross + ' ' + $Msg_Err_Sys + ' ' + $RecentError.Id) 'Red'
            }
        }
        
        # --- 異常處理流程 (紅色嚴重錯誤) ---
        if ($ErrorTriggered) {
            # 1. 計算時長 (變數供後續使用)
            $FinalDur = New-TimeSpan -Start $ScriptStartTime -End (Get-Date)
            $FinalTimeStr = "{0:D2}小時{1:D2}分鐘" -f [int][Math]::Floor($FinalDur.TotalHours), $FinalDur.Minutes

            # 2. 顯示：❌ 觸發保護
            Write-Log ($Icon_Cross + ' ' + $Msg_Prot_Trig + ' ' + $ErrorReason) 'Red' $true
            
            # 3. 顯示：➤ 觸發系統保護 (若有開啟關機)
            if ($EnableShutdown) {
                Write-Log "➤ 將執行自動關機程序" 'Yellow'
            }

            # 4. 顯示：⏱️ 本次共掛機 (您要求放在核心保護之後)
            Write-Log "⏱️ 本次共掛機：$FinalTimeStr" 'Cyan'

            # 存取崩潰截圖
            if ($ReportImages.Count -eq 0 -and $CurrentBitmap) {
                $PathCrash = Save-BitmapToFile $CurrentBitmap 'Crash'
                if ($PathCrash) { $ReportImages += $PathCrash }
            }

            # [修正] 嘗試關閉 KTK，捕捉真實錯誤訊息，不再武斷判定為權限問題
            if ($KTKProcess) { 
                try { 
                    Stop-Process -Name 'KeyToKey' -Force -ErrorAction Stop 
                } catch { 
                    # 顯示系統回傳的真實錯誤原因 (例如：存取被拒、處理程序已結束...等)
                    Write-Log "⚠️ 無法強制關閉 KeyToKey: $($_.Exception.Message)" 'Yellow'
                } 
            }
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

        # 正常運作：更新畫面緩存
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
            # 這樣做是為了確保 console 顯示的時間跟稍後 discord embed 裡讀到的日誌內容時間一致
            Write-Log ('➤ ' + $Msg_Sent_Report + " (已運行時間：$HbTimeStr)") 'Cyan'
            
            # 3. 再發送 Discord
            $HbPath = Save-BitmapToFile $Global:LastBitmapCache 'Heartbeat'
            $HbPaths = if ($HbPath) { @($HbPath) } else { @() }
            
            Send-Discord-Report -Title ($Icon_Heart + ' ' + $Msg_Discord_HB) -Reason 'Heartbeat' -ColorType 'Green' -ImagePaths $HbPaths -IsHeartbeat $true
            
            $Global:LastHeartbeatTime = Get-Date
        }

        # --- KTK 自動修復 (如果遊戲還在但腳本掛了) ---
        if ($GameProcess -and !$KTKProcess) {
            Write-Log ($Icon_Start + ' ' + $Msg_KTK_Restart) 'White' $true
            if (Test-Path $KeyToKeyPath) {
                Start-Process $KeyToKeyPath
                Write-Log $Msg_Wait_Load 'DarkGray'; Start-Sleep 35
                
                $NewKTK = Get-Process -Name 'KeyToKey' -ErrorAction SilentlyContinue
                if ($NewKTK) {
                    [Win32Tools]::ShowWindow($NewKTK.MainWindowHandle, [Win32Tools]::SW_RESTORE) | Out-Null
                    [Win32Tools]::SetForegroundWindow($NewKTK.MainWindowHandle) | Out-Null
                    Start-Sleep 1
                    Write-Log ($Msg_Send_Key + ' (' + $TargetKeyName + ')...') 'Cyan'
                    Send-Key-Native $TargetKeyName | Out-Null
                    Start-Sleep 1
                }
                Ensure-Game-TopMost
                Write-Log ($Icon_Check + ' ' + $Msg_Recovered) 'Green'
            }
        }
    }
} catch {
    Write-Host "`n[嚴重錯誤] 程式發生未預期的例外狀況：" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host "`n請截圖此畫面並回報。" -ForegroundColor Yellow
    Write-Host "按 Enter 鍵離開..." -NoNewline
    Read-Host
}