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
    $Msg_Title_Start    = '看門狗 v1.0.4 已啟動'
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
    $Msg_Stop_Monitor   = '監控已停止。按 Enter 離開視窗...'
    $Msg_Status_OK      = '掛機運作中'
    $Msg_Sent_Report    = '已發送定期 Discord 報告'
    $Msg_KTK_Restart    = 'KeyToKey 重啟中...'
    $Msg_Wait_Load      = '等待 35 秒載入...'
    $Msg_Send_Key       = '發送按鍵'
    $Msg_Recovered      = '復原完畢'
    $Msg_Footer_Base    = 'Watchdog v1.0.4'
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
        if ($Global:SessionLog.Count -gt 1200) { 
            $oldCount = $Global:SessionLog.Count
            $Global:SessionLog = $Global:SessionLog[-1000..-1] 
            $removed = $oldCount - 1000
            $Global:LastReportLogIndex = [Math]::Max(0, $Global:LastReportLogIndex - $removed)
        }
    }

    # -----------------------------------------------------------
    # 函式：Ensure-Game-TopMost
    # 功能：檢查遊戲視窗，若被縮小則還原，並嘗試將其設為前景
    # -----------------------------------------------------------
    function Ensure-Game-TopMost {
        $GameProc = Get-Process -Name 'nie' -ErrorAction SilentlyContinue
        if ($GameProc) {
            # 如果是陣列,只取第一個
            if ($GameProc -is [array]) { $GameProc = $GameProc[0] }
            
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
    # [修正] 徹底移除標題字串中的 Emoji，解決方塊亂碼問題
    # -----------------------------------------------------------
    function Show-Crash-Warning-GUI {
        param([string]$Reason)

        # 配色主題 (深色系)
        $Theme = [PSCustomObject]@{
            Bg       = [System.Drawing.Color]::FromArgb(32, 34, 37)
            Panel    = [System.Drawing.Color]::FromArgb(47, 49, 54)
            Accent   = [System.Drawing.Color]::FromArgb(240, 71, 71) # 警示紅
            TextPri  = [System.Drawing.Color]::FromArgb(255, 255, 255)
            TextSec  = [System.Drawing.Color]::FromArgb(185, 187, 190)
            BtnBg    = [System.Drawing.Color]::FromArgb(88, 101, 242)
            BtnHover = [System.Drawing.Color]::FromArgb(71, 82, 196)
        }

        # 1. 建立視窗
        $Form = New-Object System.Windows.Forms.Form
        $Form.Size = New-Object System.Drawing.Size(600, 460) # 高度微調
        $Form.StartPosition = 'CenterScreen'
        $Form.TopMost = $true
        $Form.FormBorderStyle = 'None'
        $Form.BackColor = $Theme.Bg
        $Form.ShowInTaskbar = $false

        # 2. 主面板
        $MainPanel = New-Object System.Windows.Forms.Panel
        $MainPanel.Dock = 'Fill'
        $MainPanel.BackColor = $Theme.Panel
        $Form.Controls.Add($MainPanel)

        # 3. 頂部紅色警戒條 (視覺強調，代替 Emoji)
        $TopBar = New-Object System.Windows.Forms.Panel
        $TopBar.Height = 8 # 加厚紅色條
        $TopBar.Dock = 'Top'
        $TopBar.BackColor = $Theme.Accent
        $MainPanel.Controls.Add($TopBar)

        # 輔助函式：建立標籤 (解決 if 語法相容性)
        function New-Label ($Text, $FontSize, $Bold, $Color, $Y, $Height) {
            if ($Bold) { $FontStyle = [System.Drawing.FontStyle]::Bold } else { $FontStyle = [System.Drawing.FontStyle]::Regular }
            
            $Lbl = New-Object System.Windows.Forms.Label
            $Lbl.Text = $Text
            $Lbl.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", $FontSize, $FontStyle)
            $Lbl.ForeColor = $Color
            $Lbl.AutoSize = $false
            $Lbl.Size = New-Object System.Drawing.Size(600, $Height)
            $Lbl.Location = New-Object System.Drawing.Point(0, $Y)
            $Lbl.TextAlign = 'MiddleCenter'
            return $Lbl
        }

        # [標題] 修正：移除 "⚠️"，只保留純文字
        # Y=45 (增加頂部留白)
        $LblTitle = New-Label "掛機失敗 — 關機預警" 24 $true $Theme.Accent 45 50
        $MainPanel.Controls.Add($LblTitle)

        # [原因] 
        # Y=100
        if ($Reason.Length -gt 55) { $DispReason = $Reason.Substring(0, 52) + "..." } else { $DispReason = $Reason }
        $LblReason = New-Label $DispReason 12 $false $Theme.TextSec 100 30
        $MainPanel.Controls.Add($LblReason)

        # [倒數數字] 
        # Y=140
        $LblCount = New-Object System.Windows.Forms.Label
        $LblCount.Text = "60"
        $LblCount.Font = New-Object System.Drawing.Font("Segoe UI", 75, [System.Drawing.FontStyle]::Bold)
        $LblCount.ForeColor = $Theme.TextPri
        $LblCount.AutoSize = $false
        $LblCount.Size = New-Object System.Drawing.Size(600, 120)
        $LblCount.Location = New-Object System.Drawing.Point(0, 140)
        $LblCount.TextAlign = 'MiddleCenter'
        $MainPanel.Controls.Add($LblCount)

        # [倒數說明]
        # Y=265 (拉開與數字的距離，避免擠在一起)
        $LblSub = New-Label "秒後將執行系統保護關機..." 12 $false $Theme.TextSec 265 30
        $MainPanel.Controls.Add($LblSub)

        # [裝飾進度條]
        $Prog = New-Object System.Windows.Forms.ProgressBar
        $Prog.Minimum = 0; $Prog.Maximum = 60; $Prog.Value = 60
        $Prog.Style = 'Continuous'
        $Prog.Height = 4
        $Prog.Location = New-Object System.Drawing.Point(60, 310)
        $Prog.Size = New-Object System.Drawing.Size(480, 4)
        $MainPanel.Controls.Add($Prog)

        # [取消按鈕] 
        # Y=340 (底部留白)
        $BtnCancel = New-Object System.Windows.Forms.Button
        $BtnCancel.Text = "取消關機" # 純文字，無 Emoji
        $BtnCancel.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 16, [System.Drawing.FontStyle]::Bold)
        $BtnCancel.Size = New-Object System.Drawing.Size(260, 60)
        $BtnCancel.Location = New-Object System.Drawing.Point(170, 340) # (600-260)/2 = 170
        $BtnCancel.BackColor = $Theme.BtnBg
        $BtnCancel.ForeColor = $Theme.TextPri
        $BtnCancel.FlatStyle = 'Flat'
        $BtnCancel.FlatAppearance.BorderSize = 0
        $BtnCancel.Cursor = [System.Windows.Forms.Cursors]::Hand
        $BtnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        
        $BtnCancel.Add_MouseEnter({ $BtnCancel.BackColor = $Theme.BtnHover })
        $BtnCancel.Add_MouseLeave({ $BtnCancel.BackColor = $Theme.BtnBg })
        
        $MainPanel.Controls.Add($BtnCancel)

        # 4. Timer 邏輯
        $Script:CountDown = 60
        $Timer = New-Object System.Windows.Forms.Timer
        $Timer.Interval = 1000
        $Timer.Add_Tick({
            $Script:CountDown--
            $LblCount.Text = "$Script:CountDown"
            $Prog.Value = [Math]::Max(0, [Math]::Min(60, $Script:CountDown))
            if ($Script:CountDown -le 0) {
                $Timer.Stop()
                $Form.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $Form.Close()
            }
        })
        $Timer.Start()

        # 5. 圓角效果
        $Form.Add_Shown({
            try {
                $r = 24
                $p = New-Object System.Drawing.Drawing2D.GraphicsPath
                $p.AddArc(0, 0, $r, $r, 180, 90)
                $p.AddArc($Form.Width - $r, 0, $r, $r, 270, 90)
                $p.AddArc($Form.Width - $r, $Form.Height - $r, $r, $r, 0, 90)
                $p.AddArc(0, $Form.Height - $r, $r, $r, 90, 90)
                $p.CloseFigure()
                $Form.Region = [System.Drawing.Region]::new($p)
            } catch {}
            $BtnCancel.Focus()
        })

        $Result = $Form.ShowDialog()
        $Timer.Stop()
        $Form.Dispose()
        return $Result
    }

    # -----------------------------------------------------------
    # 函式：Send-Discord-Report
    # 功能：組合訊息、日誌與截圖，發送 Embed 到 Discord
    # 優化重點：
    # 1. 使用 Using 區塊自動釋放資源 (防止檔案鎖定)
    # 2. 改善錯誤處理邏輯，避免重複嘗試刪除失敗
    # -----------------------------------------------------------
    function Send-Discord-Report {
        param(
            [string]$Title, 
            [string]$Reason, 
            [string]$ColorType='Green', 
            [string[]]$ImagePaths=@(), 
            [bool]$IsHeartbeat=$false
        )
        
        # 如果沒有設定 Webhook，直接返回不做任何事
        if ([string]::IsNullOrWhiteSpace($DiscordWebhookUrl) -or $DiscordWebhookUrl -eq 'YOUR_WEBHOOK_HERE') { 
            return 
        }
        
        # 如果不是心跳包，顯示上傳訊息
        if (!$IsHeartbeat) { 
            Write-Log 'Uploading Report...' 'Cyan' $false 
        }

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
        if ([string]::IsNullOrWhiteSpace($LogPreview)) { 
            $LogPreview = '(無)' 
        }
        
        # 將完整日誌寫入檔案
        $Global:SessionLog | Out-File -FilePath $LogSavePath -Encoding UTF8

        # --- 建立 Discord Embed 訊息 ---
        # 定義不同嚴重程度的顏色代碼
        $ColorMap = @{ 
            'Green'  = 5763719   # 綠色 (正常)
            'Red'    = 15548997  # 紅色 (嚴重錯誤)
            'Yellow' = 16705372  # 黃色 (警告)
            'Blue'   = 5793266   # 藍色 (資訊)
            'Grey'   = 9807270   # 灰色 (系統訊息)
        }
        
        # 計算程式已經運行多久
        $Duration = New-TimeSpan -Start $ScriptStartTime -End (Get-Date)
        $RunTimeStr = "{0:D2}小時{1:D2}分鐘" -f [int][Math]::Floor($Duration.TotalHours), $Duration.Minutes

        $DescHeader = ''
        $MentionContent = ''

        if ($IsHeartbeat) {
            # 心跳包：顯示系統運作正常
            $DescHeader = '**' + $Icon_Check + ' ' + $Msg_Discord_HBTxt + '**' + $LF + 
                          $Msg_Discord_SysOK + $LF + 
                          "(每 $Global:HeartbeatInterval 分鐘回報一次)" + $LF + $LF + 
                          '⏱️ **已運行時間**' + $LF + $RunTimeStr
        } else {
            # 異常通知：顯示錯誤原因與掛機時長
            $DescHeader = "**異常原因：**$LF" + $Reason + $LF + $LF + 
                          "⏳ **已掛機：**$LF" + $RunTimeStr
            
            # 紅燈與黃燈都要 @使用者 (提醒用戶注意)
            if ($ColorType -eq 'Red' -or $ColorType -eq 'Yellow') { 
                $MentionContent = "<@$DiscordUserID>" 
            }
        }

        # 如果是黃色警告，只顯示原因，不顯示日誌預覽區塊 (避免洗版)
        if ($ColorType -ne 'Yellow') {
            $EmbedDesc = $DescHeader + $LF + $LF + 
                         '**📋 ' + $Msg_Discord_Log + '**' + $LF + 
                         '```' + $LF + $LogPreview + $LF + '```'
        } else { 
            $EmbedDesc = $DescHeader 
        }

        # 建立頁尾資訊 (版本號 + 時間戳記)
        $FooterTxt = $Msg_Footer_Base + ' ' + $Icon_Bullet + ' ' + 
                     (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        
        # 組合完整的 Embed 物件
        $Embed = @{ 
            title       = $Title
            description = $EmbedDesc
            color       = $ColorMap[$ColorType]
            footer      = @{ text = $FooterTxt } 
        }
        
        # 組合最終的 Discord Payload (包含 @使用者 標記)
        $Payload = @{ 
            content = $MentionContent
            embeds  = @($Embed) 
        }
        
        # 轉為 JSON 並確保 UTF-8 編碼
        $JsonPayload = $Payload | ConvertTo-Json -Depth 10 -Compress

        # --- 發送 Multipart HTTP 請求 ---
        # 這個區塊負責把訊息、截圖、日誌檔案一起打包上傳到 Discord
        $HttpClient = $null
        $Form = $null
        $FileStreams = [System.Collections.ArrayList]::new()  # 改用 ArrayList 方便管理
        
        try {
            $HttpClient = New-Object System.Net.Http.HttpClient
            $Form = New-Object System.Net.Http.MultipartFormDataContent
            $Enc = [System.Text.Encoding]::UTF8
            
            # 加入 JSON 訊息本體
            $Form.Add(
                (New-Object System.Net.Http.StringContent($JsonPayload, $Enc, 'application/json')), 
                'payload_json'
            )

            # 加入所有截圖檔案
            $ImgIndex = 1
            foreach ($Path in $ImagePaths) {
                if (![string]::IsNullOrEmpty($Path) -and (Test-Path $Path)) {
                    $FS = [System.IO.File]::OpenRead($Path)
                    [void]$FileStreams.Add($FS)  # 記錄檔案串流，稍後統一關閉
                    
                    $ImgContent = New-Object System.Net.Http.StreamContent($FS)
                    $ImgContent.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse('image/png')
                    $Form.Add($ImgContent, "file$ImgIndex", [System.IO.Path]::GetFileName($Path))
                    $ImgIndex++
                }
            }
            
            # 僅在「非心跳」且「非黃色警告」時，才附上完整 Log 文字檔
            if (!$IsHeartbeat -and $ColorType -ne 'Yellow' -and (Test-Path $LogSavePath)) {
                $FS2 = [System.IO.File]::OpenRead($LogSavePath)
                [void]$FileStreams.Add($FS2)
                
                $TxtContent = New-Object System.Net.Http.StreamContent($FS2)
                $TxtContent.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse('text/plain')
                $Form.Add($TxtContent, "file_log", 'Watchdog_Log.txt')
            }
            
            # 加入 charset=utf-8 確保中文顯示正常
            $HttpClient.PostAsync($DiscordWebhookUrl, $Form).Result | Out-Null
            
        } catch { 
            Write-Log "Discord 上傳失敗: $_" 'Red' $true 
        } finally {
            # [優化重點] 使用 finally 確保資源一定會被釋放
            # 即使發生錯誤也能正確關閉檔案，避免檔案被鎖定
            foreach ($Stream in $FileStreams) { 
                try { $Stream.Close(); $Stream.Dispose() } catch {} 
            }
            if ($HttpClient) { try { $HttpClient.Dispose() } catch {} }
            if ($Form) { try { $Form.Dispose() } catch {} }
        }

        # [優化] 等待 1.5 秒確保檔案完全釋放後再刪除 (原本 1 秒有時候不夠)
        Start-Sleep -Milliseconds 1500
        
        # 刪除所有暫存檔案 (截圖與日誌)
        foreach ($Path in $ImagePaths) { 
            if (Test-Path $Path) { 
                try { 
                    Remove-Item $Path -Force -ErrorAction Stop 
                } catch { 
                    # 如果刪除失敗，記錄但不影響程式運作
                    Write-Log "⚠️ 無法刪除截圖: $Path" 'DarkGray'
                } 
            } 
        }
        
        if (Test-Path $LogSavePath) { 
            try { 
                Remove-Item $LogSavePath -Force -ErrorAction Stop 
            } catch { 
                Write-Log "⚠️ 無法刪除日誌檔: $LogSavePath" 'DarkGray'
            } 
        }
    }

    # -----------------------------------------------------------
    # 函式：Suppress-Steam-Window
    # 功能：偵測並關閉 Steam 干擾視窗，關閉後強制將遊戲視窗置頂
    # 優化重點：加入快取機制，避免重複偵測相同視窗
    # -----------------------------------------------------------
    # 定義全域快取變數 (記錄已關閉過的視窗，避免重複報告)
    if (-not $Global:ClosedSteamWindows) {
        $Global:ClosedSteamWindows = @{}  # 用 Hashtable 儲存 (視窗標題 -> 最後關閉時間)
    }
    
    function Suppress-Steam-Window {
        # 使用 C# 核心的 FindSteamWindows 進行全域搜尋
        $Targets = [SteamBuster]::FindSteamWindows()
        $ClosedAny = $false
        $Now = Get-Date
        
        foreach ($win in $Targets) {
            # [優化] 檢查這個視窗是否在 30 秒內已經關閉過
            # 如果是，跳過不處理 (避免重複報告)
            $CacheKey = $win.Title + "|" + $win.ProcessName
            if ($Global:ClosedSteamWindows.ContainsKey($CacheKey)) {
                $LastClosed = $Global:ClosedSteamWindows[$CacheKey]
                $TimeDiff = ($Now - $LastClosed).TotalSeconds
                if ($TimeDiff -lt 30) {
                    continue  # 跳過這個視窗
                }
            }
            
            # 1. 偵測到干擾：寫 Log + 顯示
            $Msg = "偵測到干擾視窗！標題: [$($win.Title)] (程式: $($win.ProcessName))"
            Write-Log ($Icon_Warn + ' ' + $Msg) 'Yellow' $true
            
            # 2. 發送警告
            Send-Discord-Report -Title ($Icon_Warn + ' 異常徵兆警告') `
                               -Reason "$Msg`n(已執行自動關閉)" `
                               -ColorType 'Yellow'
            
            # 3. 執行關閉
            [SteamBuster]::CloseWindow($win.Handle)
            $ClosedAny = $true
            
            # 4. 更新快取 (記錄這個視窗已經被關閉過)
            $Global:ClosedSteamWindows[$CacheKey] = $Now
            
            # 5. 寫 Log
            Write-Log ($Icon_Check + ' 已關閉視窗。') 'Green'
        }
        
        # 如果有關閉過視窗，立刻把焦點搶回遊戲，避免 KTK 按錯地方
        if ($ClosedAny) { 
            Ensure-Game-TopMost 
        }
        
        # [優化] 定期清理舊的快取紀錄 (超過 5 分鐘的項目)
        # 避免 Hashtable 無限增長佔用記憶體
        $KeysToRemove = @()
        foreach ($Key in $Global:ClosedSteamWindows.Keys) {
            $LastTime = $Global:ClosedSteamWindows[$Key]
            if (($Now - $LastTime).TotalMinutes -gt 5) {
                $KeysToRemove += $Key
            }
        }
        foreach ($Key in $KeysToRemove) {
            $Global:ClosedSteamWindows.Remove($Key)
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
        } catch { 
            return $null 
        }
    }

    # -----------------------------------------------------------
    # 函式：Get-PixelsFromBitmap
    # 功能：將高解析度 Bitmap 縮放為 12x12 的縮圖，並提取像素顏色值
    # 目的：透過極低解析度比對，大幅降低 CPU 運算量，同時忽略細微雜訊
    # 優化重點：使用 LockBits 加速像素讀取 (比 GetPixel 快 10~20 倍)
    # -----------------------------------------------------------
    function Get-PixelsFromBitmap ($Bitmap) {
        if (!$Bitmap) { return $null }
        
        try {
            # 1. 先縮放成 12x12 的小圖
            $Small = $Bitmap.GetThumbnailImage(12, 12, $null, [IntPtr]::Zero)
            
            # 2. [優化] 使用 LockBits 快速讀取像素資料
            # 傳統 GetPixel() 方法很慢，因為每次都要做很多安全檢查
            # LockBits 直接讀取記憶體，速度快很多
            $Rect = New-Object System.Drawing.Rectangle(0, 0, 12, 12)
            $BmpData = $Small.LockBits($Rect, 
                                       [System.Drawing.Imaging.ImageLockMode]::ReadOnly, 
                                       $Small.PixelFormat)
            
            $Pixels = New-Object 'int[,]' 12, 12
            $Ptr = $BmpData.Scan0
            $Stride = $BmpData.Stride
            
            # 3. 根據像素格式決定讀取方式
            if ($Small.PixelFormat -eq [System.Drawing.Imaging.PixelFormat]::Format32bppArgb) {
                # 每個像素佔 4 bytes (ARGB)
                for ($y = 0; $y -lt 12; $y++) {
                    for ($x = 0; $x -lt 12; $x++) {
                        $Offset = $y * $Stride + $x * 4
                        $Color = [System.Runtime.InteropServices.Marshal]::ReadInt32($Ptr, $Offset)
                        $Pixels[$x, $y] = $Color
                    }
                }
            } else {
                # 其他格式退回慢速方法
                $Small.UnlockBits($BmpData)
                for ($x = 0; $x -lt 12; $x++) { 
                    for ($y = 0; $y -lt 12; $y++) { 
                        $Pixels[$x, $y] = $Small.GetPixel($x, $y).ToArgb() 
                    } 
                }
                $Small.Dispose()
                return ,$Pixels
            }
            
            $Small.UnlockBits($BmpData)
            $Small.Dispose()
            return ,$Pixels
            
        } catch {
            # 發生錯誤時回傳 null
            return $null
        }
    }

    # -----------------------------------------------------------
    # 函式：Get-Similarity
    # 功能：比對兩組像素矩陣的相似度 (0~100%)
    # -----------------------------------------------------------
    function Get-Similarity ($PixA, $PixB) {
        if (!$PixA -or !$PixB) { return 0 }
        
        $Match = 0
        $Total = 144  # 12x12 = 144 個像素
        
        for ($x = 0; $x -lt 12; $x++) { 
            for ($y = 0; $y -lt 12; $y++) {
                $valA = $PixA[$x, $y]
                $valB = $PixB[$x, $y]
                
                # 分離 RGB 三個顏色通道
                $R1 = ($valA -shr 16) -band 255
                $G1 = ($valA -shr 8) -band 255
                $B1 = $valA -band 255
                
                $R2 = ($valB -shr 16) -band 255
                $G2 = ($valB -shr 8) -band 255
                $B2 = $valB -band 255
                
                # 如果三個通道的差異都小於 20，視為相同
                if ([Math]::Abs($R1 - $R2) -lt 20 -and 
                    [Math]::Abs($G1 - $G2) -lt 20 -and 
                    [Math]::Abs($B1 - $B2) -lt 20) { 
                    $Match++ 
                }
            }
        }
        
        # 計算相似度百分比，四捨五入到小數點後一位
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
        
        try {
            $Bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
            return $Path
        } catch {
            Write-Log "⚠️ 無法儲存截圖: $_" 'Yellow'
            return $null
        }
    }

    # ==========================================
    # 4. Initialization（初始化流程）
    # ==========================================
    Clear-Host
    try { [Console]::CursorVisible = $true } catch {}
    
    Write-Host '==========================================' -ForegroundColor Cyan
    Write-Host '   Victory Road & KeyToKey Watchdog v1.0.4' -ForegroundColor Cyan
    Write-Host '   (優化版本 - 2024)' -ForegroundColor Cyan
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
    
    if ($EnableShutdown) { 
        Write-Host 'Y (已啟用關機保護)' -ForegroundColor Red 
    } else { 
        Write-Host 'N (僅關閉程式)' -ForegroundColor Green 
    }

    # 詢問按鍵設定
    Write-Host ''
    Write-Host '[設定] 請輸入 KTK 啟動熱鍵  [預設: F7]' -ForegroundColor Yellow
    $InputKey = Read-Host '請輸入'
    if ([string]::IsNullOrWhiteSpace($InputKey)) { $TargetKeyName = 'F7' } else { $TargetKeyName = $InputKey.Trim().ToUpper() }
    Write-Host ('已設定按鍵: ' + $TargetKeyName) -ForegroundColor Green

    # 詢問心跳頻率
    Write-Host ''
    Write-Host '[設定] 請輸入 Discord 定期回報間隔 (分鐘) [預設: 5 分鐘]' -ForegroundColor Yellow
    $InputInterval = Read-Host '請輸入'
    if (![string]::IsNullOrWhiteSpace($InputInterval) -and ($InputInterval -match '^\d+$')) {
        $v = [int]$InputInterval
        # 限制範圍: 最少 1 分鐘,最多 1440 分鐘 (24 小時)
        $Global:HeartbeatInterval = [Math]::Max(1, [Math]::Min($v, 1440))
    }
    Write-Host ('已設定回報間隔: ' + $Global:HeartbeatInterval + ' 分鐘') -ForegroundColor Green

    try { [Console]::CursorVisible = $false } catch {}

    # ==========================================
    # 5. Main Loop（主監控迴圈）
    # ==========================================
    $FreezeCount = 0
    $NoResponseCount = 0
    
    # 監控開始前先截一張圖作為基準
    $Global:LastBitmapCache = Capture-ScreenBitmap
    $LastPixelData = Get-PixelsFromBitmap $Global:LastBitmapCache

    Write-Host ''
    Write-Host '=== 監控開始 (按 Q 停止並回報) ===' -ForegroundColor Cyan
    
    # 手動清空一次日誌，確保監控開始前的雜訊不會被計入
    $Global:SessionLog = @()
    Send-Discord-Report -Title ($Icon_Start + ' ' + $Msg_Title_Start) `
                       -Reason $Msg_Reason_Start `
                       -ColorType 'Blue' `
                       -IsHeartbeat $true

    # [初始化檢查] 啟動前環境檢查 (不等待倒數)
    Ensure-Game-TopMost
    
    # 檢查是否需要啟動 KeyToKey
    if ((Get-Process -Name 'nie' -ErrorAction SilentlyContinue) -and 
        !(Get-Process -Name 'KeyToKey' -ErrorAction SilentlyContinue)) {
        
        Write-Log "➤ 初始檢查：KeyToKey 未執行，嘗試啟動..." 'Yellow'
        
        if (Test-Path $KeyToKeyPath) { 
            Start-Process $KeyToKeyPath
            Write-Log $Msg_Wait_Load 'DarkGray'
            Start-Sleep 35
            
            # 抓取視窗並按下熱鍵
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
        }
    }

    # ========== 主要監控迴圈開始 ==========
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
                    
                    Send-Discord-Report -Title ($Icon_Stop + ' ' + $Msg_Title_Stop) `
                                       -Reason "$Msg_Reason_Stop`n⏱️ **共運行：**$FinalTimeStr" `
                                       -ColorType 'Grey'
                    
                    Write-Host $Msg_Stop_Monitor -ForegroundColor Green
                    Read-Host
                    exit
                }
            }

            # 2. [每 5 秒執行] 檢查遊戲與 KTK 運作狀態
            # [優化原理] Get-Process 指令比較耗時，改成每 5 秒查一次
            if ($i % 5 -eq 0) {
                $CheckGame = Get-Process -Name 'nie' -ErrorAction SilentlyContinue
                $CheckKTK = Get-Process -Name 'KeyToKey' -ErrorAction SilentlyContinue
                $StatusStr = ''
                
                if ($CheckGame) { 
                    if ($CheckGame.Responding) { 
                        $StatusStr += $Msg_Game_Run 
                    } else { 
                        $StatusStr += $Msg_Game_NoResp + ' ' + $Icon_Warn 
                    } 
                } else { 
                    $StatusStr += $Msg_Game_Lost + ' ' + $Icon_Cross 
                }
                
                $StatusStr += ' | '
                
                if ($CheckKTK) { 
                    $StatusStr += $Msg_KTK_Run 
                } else { 
                    $StatusStr += $Msg_KTK_Err + ' ' + $Icon_Warn 
                }
            }

            # 計算進度條百分比
            if ($LoopIntervalSeconds -gt 1) { 
                $Percent = ($LoopIntervalSeconds - $i) / ($LoopIntervalSeconds - 1) 
            } else { 
                $Percent = 1 
            }
            $ProgressCount = [int][Math]::Floor($Percent * 20)
            
            # 強制游標起步至少為 1
            if ($ProgressCount -lt 1) { $ProgressCount = 1 }
            
            # 游標閃爍 (每秒閃爍兩次)
            for ($blink = 0; $blink -lt 2; $blink++) {
                $BarStr = ''
                
                if ($ProgressCount -ge 20) { 
                    $BarStr = '=' * 20 
                } else { 
                    if ($ProgressCount -gt 0) { 
                        if ($blink -eq 0) { 
                            $BarStr = '=' * $ProgressCount 
                        } else { 
                            $BarStr = '=' * ($ProgressCount - 1) + ' ' 
                        } 
                    } 
                }
                
                $Bar = '[' + $BarStr + (' ' * (20 - $BarStr.Length)) + ']'
                
                # 使用歸位字元覆蓋同一行文字
                Write-Host ($CR + $Bar + " 倒數 $i 秒... (按 Q 停止) [ $StatusStr ]       ") `
                          -NoNewline -ForegroundColor Gray
                
                Start-Sleep -Milliseconds 500
            }
        }
        Write-Host ''  # 倒數結束換行

        # --- 核心檢測邏輯 ---
        $GameProcess = Get-Process -Name 'nie' -ErrorAction SilentlyContinue
        $KTKProcess = Get-Process -Name 'KeyToKey' -ErrorAction SilentlyContinue
        $ErrorTriggered = $false
        $ErrorReason = ''
        
        # 準備變數：擷取當前畫面
        $CurrentBitmap = Capture-ScreenBitmap
        $CurrentPixelData = Get-PixelsFromBitmap $CurrentBitmap
        $ReportImages = @()

        # ==========================================================================
        # 1. [核心檢測] 程式消失 (Crash)
        # ==========================================================================
        if (!$GameProcess) { 
            $ErrorTriggered = $true
            
            # [設定範圍] 只搜尋過去 5 分鐘內的錯誤
            $TimeLimit = (Get-Date).AddMinutes(-5) 
            
            # [極速優化] 使用 StartTime 參數直接篩選
            $SysErrs = Get-WinEvent -FilterHashtable @{
                LogName='System'
                Id=141,4101,117
                StartTime=$TimeLimit
            } -ErrorAction SilentlyContinue
            
            $AppErrs = Get-WinEvent -FilterHashtable @{
                LogName='Application'
                Id=1001
                StartTime=$TimeLimit
            } -ErrorAction SilentlyContinue | Where-Object { 
                $_.Message -match 'LiveKernelEvent' 
            }
            
            # 合併並排序
            $AllErrs = @($SysErrs) + @($AppErrs) | Sort-Object TimeCreated -Descending
            
            if ($AllErrs) {
                # 有找到錯誤
                $RecentError = $AllErrs | Select-Object -First 1
                
                # 代碼解析
                if ($RecentError.Id -eq 1001) {
                    if ($RecentError.Message -match '141') { 
                        $ErrCode = "LiveKernelEvent (141)" 
                    } elseif ($RecentError.Message -match '117') { 
                        $ErrCode = "LiveKernelEvent (117)" 
                    } elseif ($RecentError.Message -match '1a1') { 
                        $ErrCode = "LiveKernelEvent (1a1)" 
                    } else { 
                        $ErrCode = "LiveKernelEvent (1001)" 
                    }
                } else {
                    $ErrCode = $RecentError.Id
                }
                
                $SysErrMsg = $Msg_Err_Reason + $ErrCode + ')'
                $ErrorReason = $SysErrMsg
                
                Write-Log ($Icon_Cross + ' 偵測到程式消失，並發現系統錯誤: ' + $ErrCode) 'Red'
            } else { 
                # 沒找到錯誤
                $ErrorReason = $Msg_Err_Crash 
            }
        }

        # 2. 檢測：無回應
        if ($GameProcess -and !$GameProcess.Responding) {
            $NoResponseCount++
            Write-Log ($Icon_Warn + ' ' + $Msg_Warn_NoResp + ' (' + $NoResponseCount + '/' + $NoResponseThreshold + ')') 'Yellow'
            
            if ($NoResponseCount -ge $NoResponseThreshold) { 
                $ErrorTriggered = $true
                $ErrorReason = $Msg_Err_NoResp
                Stop-Process -Name 'nie' -Force -ErrorAction SilentlyContinue 
            }
        } else { 
            $NoResponseCount = 0 
        }

        # 3. 檢測：凍結
        if ($CurrentPixelData -and $LastPixelData) {
            $Similarity = Get-Similarity $CurrentPixelData $LastPixelData
            
            if ($Similarity -ge $FreezeSimilarity) {
                $FreezeCount++
                Write-Log ($Icon_Warn + ' ' + $Msg_Warn_Freeze + $Similarity + '%) (' + $FreezeCount + '/' + $FreezeThreshold + ')') 'Yellow'
                
                # 存圖保留證據
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
                    # 達到閾值，判定為死當
                    $ErrorTriggered = $true
                    $ErrorReason = $Msg_Err_Freeze 
                } else { 
                    # 未達閾值 (例如只凍結 1 次或 2 次)，發送黃色警告給 Discord
                    # 這樣使用者可以提早知道「遊戲好像有點怪怪的」，但還不到需要關機的程度
                    Send-Discord-Report -Title ($Icon_Warn + ' ' + $Msg_Discord_Title_W) `
                                       -Reason "畫面相似度過高 ($Similarity%) - 累積 $FreezeCount/$FreezeThreshold" `
                                       -ColorType 'Yellow' `
                                       -ImagePaths $ReportImages 
                }
            } else { 
                # 畫面有變化，代表遊戲正常運作中，重置凍結計數器
                $FreezeCount = 0 
            }
        }

        # ==========================================================================
        # 4. [輔助檢測] 系統錯誤 (被動查驗模式)
        #    邏輯說明：只有當「前面已經判定有問題」時，才執行這段。
        #    目的：補充更詳細的錯誤資訊到報告中，但絕不主動殺遊戲。
        # ==========================================================================
        if ($ErrorTriggered) {
            # 設定時間範圍：只查過去 5 分鐘內的系統日誌
            $TimeLimit = (Get-Date).AddMinutes(-5)
            
            # [查詢 System 日誌] 尋找常見的硬體錯誤代碼
            # - 141: 顯示卡驅動程式停止回應並已復原 (最常見的遊戲崩潰原因)
            # - 4101: 顯示器驅動程式已停止回應 (類似 141)
            # - 117: 顯示卡超時偵測與復原 (TDR)
            $SysErrs = Get-WinEvent -FilterHashtable @{
                LogName='System'           # 系統日誌
                Id=141,4101,117           # 要找的錯誤代碼
                StartTime=$TimeLimit      # 只看最近 5 分鐘
            } -ErrorAction SilentlyContinue  # 如果沒找到也不報錯
            
            # [查詢 Application 日誌] 尋找 Windows 錯誤報告 (WER) 的 LiveKernelEvent
            # 有時候 141 錯誤會被包裝在 1001 事件裡面，需要額外解析訊息內容
            $AppErrs = Get-WinEvent -FilterHashtable @{
                LogName='Application'     # 應用程式日誌
                Id=1001                   # Windows 錯誤報告事件
                StartTime=$TimeLimit      # 只看最近 5 分鐘
            } -ErrorAction SilentlyContinue | Where-Object { 
                # 進一步過濾：訊息內容必須包含 'LiveKernelEvent' 才算數
                $_.Message -match 'LiveKernelEvent' 
            }
            
            # 將兩邊找到的錯誤合併成一個陣列，並按時間倒序排列 (最新的在最前面)
            $AllErrs = @($SysErrs) + @($AppErrs) | Sort-Object TimeCreated -Descending
            
            if ($AllErrs) {
                # 有找到系統錯誤！抓出最新的那一筆
                $RecentError = $AllErrs | Select-Object -First 1
                
                # [錯誤代碼解析] 判斷這個錯誤到底是哪一種
                if ($RecentError.Id -eq 1001) {
                    # 如果是 1001 (WER 事件)，需要進一步解析訊息內容
                    # 因為真正的錯誤代碼被包在訊息字串裡面
                    if ($RecentError.Message -match '141') { 
                        $ErrCode = "LiveKernelEvent (141)"  # 顯卡驅動崩潰
                    } elseif ($RecentError.Message -match '117') { 
                        $ErrCode = "LiveKernelEvent (117)"  # 顯卡超時
                    } elseif ($RecentError.Message -match '1a1') { 
                        $ErrCode = "LiveKernelEvent (1a1)"  # 顯卡硬體錯誤
                    } else { 
                        $ErrCode = "LiveKernelEvent (1001)" # 其他未知的 LiveKernelEvent
                    }
                } else {
                    # 如果是 141、4101、117 這些直接的 System 事件，ID 就是錯誤代碼
                    $ErrCode = $RecentError.Id
                }
                
                # 組合完整的系統錯誤訊息
                $SysErrMsg = $Msg_Err_Reason + $ErrCode + ')'
                
                # [靜默補充機制] 把系統錯誤資訊「偷偷」加到 $ErrorReason 變數裡
                # 為什麼叫「靜默」？因為我們不用 Write-Log 寫到畫面上，避免畫面被洗版
                # 只在 Discord 報告裡面顯示就好
                # 
                # [防重複機制] 先檢查 $ErrorReason 裡面是不是已經有這個錯誤代碼了
                # 使用 [regex]::Escape() 確保特殊字元 (例如括號) 不會被誤判成正規表達式
                if ($ErrorReason -notmatch [regex]::Escape($ErrCode)) {
                    # 如果還沒有，就把它加進去 (用換行隔開)
                    $ErrorReason += "`n[系統紀錄] $SysErrMsg"
                }
                # 如果已經有了，什麼都不做，避免重複顯示相同的錯誤
            }
        }
        
        # ==========================================================================
        # 異常處理流程：當上面的檢測邏輯判定「有問題」時，執行這段
        # ==========================================================================
        if ($ErrorTriggered) {
            # ----------------------------------------------------------------------
            # 步驟 1：計算本次掛機的總時長
            # ----------------------------------------------------------------------
            # 用「現在的時間」減去「腳本開始的時間」，得到時間差
            $FinalDur = New-TimeSpan -Start $ScriptStartTime -End (Get-Date)
            # 將時間差格式化成「XX小時XX分鐘」的字串，方便顯示
            # [int][Math]::Floor() 的作用是「無條件捨去」，例如 2.8 小時 -> 2 小時
            $FinalTimeStr = "{0:D2}小時{1:D2}分鐘" -f [int][Math]::Floor($FinalDur.TotalHours), $FinalDur.Minutes

            # ----------------------------------------------------------------------
            # 步驟 2：顯示「觸發保護」訊息到控制台與日誌
            # ----------------------------------------------------------------------
            # 使用紅色文字顯示嚴重錯誤，讓使用者一眼就能看到出了什麼問題
            Write-Log ($Icon_Cross + ' ' + $Msg_Prot_Trig + $ErrorReason) 'Red'
            
            # ----------------------------------------------------------------------
            # 步驟 3：如果有啟用「自動關機保護」，顯示警告訊息
            # ----------------------------------------------------------------------
            if ($EnableShutdown) { 
                # 黃色警告：告訴使用者「等一下電腦會關機」
                Write-Log "➤ 將執行自動關機程序" 'Yellow' 
            }
            
            # ----------------------------------------------------------------------
            # 步驟 4：顯示本次掛機的總時長
            # ----------------------------------------------------------------------
            # 用青色 (Cyan) 顯示資訊性訊息，告訴使用者這次總共掛了多久
            Write-Log "⏱️ 本次共掛機：$FinalTimeStr" 'Cyan'

            # ----------------------------------------------------------------------
            # 步驟 5：儲存崩潰現場的截圖 (如果前面還沒存過)
            # ----------------------------------------------------------------------
            # 檢查：如果 $ReportImages 陣列是空的 (代表前面的凍結檢測沒有存圖)
            # 並且當前畫面的 Bitmap 物件存在，就趕快存一張「崩潰截圖」
            if ($ReportImages.Count -eq 0 -and $CurrentBitmap) {
                # 呼叫 Save-BitmapToFile 函式，將畫面存成 PNG 檔案
                # 檔名前綴會是 'Crash_20240115_123456.png' 這種格式
                $PathCrash = Save-BitmapToFile $CurrentBitmap 'Crash'
                # 如果存檔成功 (函式回傳了檔案路徑)，就把路徑加到陣列裡，準備發送到 Discord
                if ($PathCrash) { $ReportImages += $PathCrash }
            }
            
            # ----------------------------------------------------------------------
            # 步驟 6：強制關閉 KeyToKey 程式
            # ----------------------------------------------------------------------
            # 如果 KeyToKey 程式還在執行，就把它關掉
            if ($KTKProcess) { 
                try { 
                    # 使用 Stop-Process 強制終止 (Force 參數代表「不管它在幹嘛，直接殺掉」)
                    Stop-Process -Name 'KeyToKey' -Force -ErrorAction Stop 
                } catch { 
                    # 如果關閉失敗 (例如程式卡死到連 Stop-Process 都殺不掉)
                    # 記錄錯誤訊息到日誌，但不影響後續流程
                    # $_.Exception.Message 會顯示詳細的錯誤原因 (例如「拒絕存取」)
                    Write-Log "⚠️ 無法強制關閉 KeyToKey: $($_.Exception.Message)" 'Yellow' 
                } 
            }
            
            # ----------------------------------------------------------------------
            # 步驟 7：強制關閉遊戲程式 (勝利之路)
            # ----------------------------------------------------------------------
            # 無論如何都嘗試關閉，即使失敗也不報錯 (SilentlyContinue)
            Stop-Process -Name 'nie' -Force -ErrorAction SilentlyContinue
            
            # ----------------------------------------------------------------------
            # 步驟 8：組合 Discord 通知的訊息內容
            # ----------------------------------------------------------------------
            # 如果有啟用關機功能，在錯誤原因後面加上「(已執行自動關機程序)」
            # 如果沒有啟用，就只顯示錯誤原因本身
            if ($EnableShutdown) { 
                $DiscordReason = "$ErrorReason`n(已執行自動關機程序)" 
            } else { 
                $DiscordReason = $ErrorReason 
            }
            
            # ----------------------------------------------------------------------
            # 步驟 9：發送 Discord 緊急通知
            # ----------------------------------------------------------------------
            # 呼叫 Send-Discord-Report 函式，將所有資訊打包發送
            # 參數說明：
            # - Title: 標題 (會顯示「❌ 嚴重異常終止」)
            # - Reason: 錯誤原因 (包含系統錯誤代碼、掛機時長等資訊)
            # - ColorType: 'Red' 代表紅色警報 (最高等級)
            # - ImagePaths: 截圖檔案路徑陣列 (可能包含凍結對比圖或崩潰截圖)
            Send-Discord-Report -Title ($Icon_Cross + ' ' + $Msg_Discord_Title) `
                               -Reason $DiscordReason `
                               -ColorType 'Red' `
                               -ImagePaths $ReportImages
            
            # ----------------------------------------------------------------------
            # 步驟 10：釋放記憶體中的 Bitmap 物件 (防止記憶體洩漏)
            # ----------------------------------------------------------------------
            # Bitmap 物件佔用記憶體很大 (可能好幾 MB)，用完一定要手動釋放
            # 否則會導致記憶體一直被佔用，無法歸還給系統
            if ($Global:LastBitmapCache) { 
                try { 
                    # 呼叫 .Dispose() 方法釋放資源
                    $Global:LastBitmapCache.Dispose() 
                } catch { 
                    # 如果釋放失敗，忽略錯誤 (通常不會失敗)
                } 
            }
            if ($CurrentBitmap) { 
                try { 
                    $CurrentBitmap.Dispose() 
                } catch {} 
            }

            # ----------------------------------------------------------------------
            # 步驟 11：執行關機流程 (如果使用者有啟用的話)
            # ----------------------------------------------------------------------
            if ($EnableShutdown) { 
                # [顯示 GUI 警告視窗] 倒數 60 秒，讓使用者有機會取消
                # Show-Crash-Warning-GUI 會顯示一個彈出視窗，上面有大大的倒數數字
                # 使用者可以按「取消關機」按鈕來阻止電腦關機
                $GuiResult = Show-Crash-Warning-GUI -Reason $ErrorReason
                
                # [判斷使用者的選擇]
                if ($GuiResult -eq [System.Windows.Forms.DialogResult]::OK) {
                    # 如果倒數結束 (使用者沒有按取消)，執行關機
                    Write-Log $Msg_Shutdown 'Red'
                    # Stop-Computer -Force: 強制關機，不詢問也不等待
                    Stop-Computer -Force
                    exit  # 關機指令發出後，立刻結束腳本
                } else {
                    # 如果使用者按了「取消關機」按鈕
                    Write-Log "使用者已取消關機。" 'Yellow'
                    Write-Log $Msg_Stop_Monitor 'Red'
                    Read-Host  # 等待使用者按 Enter 鍵
                    exit       # 結束腳本
                }
            } else {
                # [沒有啟用關機功能] 只關閉程式，不關機
                Write-Log $Msg_Stop_Monitor 'Red'
                Read-Host  # 等待使用者按 Enter 鍵確認已讀訊息
                exit       # 結束腳本
            }
        }

        # ==========================================================================
        # 記憶體管理：更新「上一張畫面」的快取
        # ==========================================================================
        # [原理說明] 凍結檢測的運作方式是「比對兩張畫面」
        # - 第一輪：LastBitmap 是基準，CurrentBitmap 是新畫面
        # - 第二輪：把 Current 變成 Last，下次抓新的 Current 來比對
        # 這樣就能持續監控畫面是否有變化
        
        # 先釋放舊的 LastBitmap 物件 (避免記憶體洩漏)
        if ($Global:LastBitmapCache) { 
            try { 
                $Global:LastBitmapCache.Dispose() 
            } catch { 
                # 釋放失敗也沒關係，忽略錯誤
            } 
        }
        
        # 把「當前畫面」變成「上一張畫面」，供下次迴圈使用
        # 注意：這裡是「轉移所有權」，不是複製！
        # 所以上面不能對 $CurrentBitmap 執行 Dispose()，否則下次迴圈會出錯
        $Global:LastBitmapCache = $CurrentBitmap
        
        # 同時也把像素資料更新
        $LastPixelData = $CurrentPixelData
        
        # ==========================================================================
        # 狀態顯示：如果一切正常，顯示「掛機運作中」
        # ==========================================================================
        # 條件判斷：沒有觸發異常 (ErrorTriggered = false) 且 KeyToKey 正在執行
        if (!$ErrorTriggered -and $KTKProcess) { 
            # 使用深灰色 (DarkGray) 顯示「不太重要的資訊」，避免畫面太雜亂
            Write-Log ('➤ ' + $Msg_Status_OK) 'DarkGray' 
        }

        # ==========================================================================
        # 定期心跳 (Heartbeat)：每隔 N 分鐘發送一次「我還活著」的訊息到 Discord
        # ==========================================================================
        # [功能目的] 讓使用者知道看門狗程式沒有當機，持續在監控中
        # [觸發條件] 距離上次發送心跳包已經過了 N 分鐘 (預設 5 分鐘)
        
        # 計算「現在」距離「上次發送心跳」過了多久
        $TimeSinceLastHeartbeat = (Get-Date) - $Global:LastHeartbeatTime
        
        # 判斷是否已經超過設定的間隔時間
        if ($TimeSinceLastHeartbeat.TotalMinutes -ge $Global:HeartbeatInterval) {
            # ----------------------------------------------------------------------
            # 步驟 1：計算總運行時長
            # ----------------------------------------------------------------------
            $HbDur = New-TimeSpan -Start $ScriptStartTime -End (Get-Date)
            # 格式化成「XX小時XX分鐘」
            $HbTimeStr = "{0:D2}小時{1:D2}分鐘" -f [int][Math]::Floor($HbDur.TotalHours), $HbDur.Minutes
            
            # ----------------------------------------------------------------------
            # 步驟 2：先寫入本地日誌 (Console 畫面 + 記憶體 Log)
            # ----------------------------------------------------------------------
            # 使用青色 (Cyan) 顯示資訊性訊息
            Write-Log ('➤ ' + $Msg_Sent_Report + " (已運行時間：$HbTimeStr)") 'Cyan'
            
            # ----------------------------------------------------------------------
            # 步驟 3：儲存當前畫面截圖 (證明程式真的有在跑)
            # ----------------------------------------------------------------------
            # 呼叫 Save-BitmapToFile，檔名前綴會是 'Heartbeat_20240115_123456.png'
            $HbPath = Save-BitmapToFile $Global:LastBitmapCache 'Heartbeat'
            # 如果存檔成功，把路徑包成陣列；如果失敗，用空陣列
            if ($HbPath) { $HbPaths = @($HbPath) } else { $HbPaths = @() }
            
            # ----------------------------------------------------------------------
            # 步驟 4：發送心跳包到 Discord
            # ----------------------------------------------------------------------
            # 參數說明：
            # - Title: 標題會顯示「🫀 看門狗定期報告」
            # - Reason: 這裡只是個佔位符，實際內容由 Send-Discord-Report 自動組合
            # - ColorType: 'Green' 代表綠燈 (一切正常)
            # - ImagePaths: 附上當前畫面截圖
            # - IsHeartbeat: $true 告訴函式「這是心跳包」，會使用特殊的訊息格式
            Send-Discord-Report -Title ($Icon_Heart + ' ' + $Msg_Discord_HB) `
                               -Reason $Msg_Discord_HBTxt `
                               -ColorType 'Green' `
                               -ImagePaths $HbPaths `
                               -IsHeartbeat $true
            
            # 步驟 5：更新上次心跳時間，避免下一輪迴圈重複發送
            $Global:LastHeartbeatTime = Get-Date
        }

        # ==========================================================================
        # KTK 自動修復：如果遊戲還在執行，但 KeyToKey 掛了，自動重啟它
        # ==========================================================================
        # [觸發條件] 遊戲程式 (nie) 存在 且 KeyToKey 程式不存在
        # [執行時機] 每 75 秒檢查一次（在每輪監控結束時）
        # [目的] 確保掛機腳本持續運作，不會因為 KeyToKey 崩潰而中斷
        if ($GameProcess -and !$KTKProcess) {
            # ----------------------------------------------------------------------
            # 步驟 1：顯示「正在重啟 KeyToKey」的訊息
            # ----------------------------------------------------------------------
            # 使用白色文字，並強制換行 (ForceNewLine = $true)
            # 為什麼要強制換行？因為前面可能正在顯示倒數計時，需要先換行才不會覆蓋文字
            Write-Log ('➤ ' + $Msg_KTK_Restart) 'White' $true
            
            # ----------------------------------------------------------------------
            # 步驟 2：檢查 KeyToKey 執行檔是否存在
            # ----------------------------------------------------------------------
            if (Test-Path $KeyToKeyPath) {
                # [2-1] 啟動 KeyToKey 程式
                # Start-Process 會在背景執行程式，不會阻塞腳本
                Start-Process $KeyToKeyPath
                
                # [2-2] 等待程式載入
                # 顯示「等待 35 秒載入...」訊息 (用深灰色，代表這是背景作業)
                Write-Log $Msg_Wait_Load 'DarkGray'
                # 實際暫停 35 秒，讓 KeyToKey 有足夠時間完全啟動
                # 為什麼要等這麼久？因為 KeyToKey 需要初始化介面、載入設定檔等等
                Start-Sleep 35
                
                # [2-3] 重新抓取 KeyToKey 的 Process 物件
                # 因為剛才用 Start-Process 啟動的程式，還沒有被儲存到變數裡
                # 所以需要重新用 Get-Process 去系統中找它
                $NewKTK = Get-Process -Name 'KeyToKey' -ErrorAction SilentlyContinue
                
                # [2-4] 如果成功找到 KeyToKey 程式
                if ($NewKTK) {
                    # 處理可能的陣列情況（雖然通常只有一個，但為了保險）
                    if ($NewKTK -is [array]) { $NewKTK = $NewKTK[0] }
                    
                    # 先把視窗還原 (如果它被最小化的話)
                    # SW_RESTORE = 9，代表「還原視窗到正常大小」
                    [Win32Tools]::ShowWindow($NewKTK.MainWindowHandle, [Win32Tools]::SW_RESTORE) | Out-Null
                    
                    # 把視窗設為前景 (active window)，確保它在最上層
                    # 這樣下一步發送按鍵時，才能正確送到 KeyToKey 視窗
                    [Win32Tools]::SetForegroundWindow($NewKTK.MainWindowHandle) | Out-Null
                    
                    # 等待 1 秒，確保視窗已經完全切換到前景
                    Start-Sleep 1
                    
                    # [2-5] 發送啟動熱鍵 (例如 F7)
                    # 顯示訊息：「發送按鍵 (F7)...」
                    Write-Log ($Msg_Send_Key + ' (' + $TargetKeyName + ')...') 'Cyan'
                    
                    # 呼叫 Send-Key-Native 函式，模擬實體鍵盤按下按鍵
                    # 這個函式會使用 Windows API 的 keybd_event，比 SendKeys 更可靠
                    Send-Key-Native $TargetKeyName | Out-Null
                    
                    # 等待 1 秒，讓 KeyToKey 處理按鍵事件
                    Start-Sleep 1
                }
                
                # [2-6] 確保遊戲視窗回到最上層
                # 因為剛才把焦點切換到 KeyToKey，現在要把焦點搶回遊戲
                # 這樣 KeyToKey 的按鍵指令才能正確送到遊戲裡
                Ensure-Game-TopMost
                
                # [2-7] 顯示「復原完畢」訊息
                # 使用綠色 + 打勾符號，代表修復成功
                Write-Log ($Icon_Check + ' ' + $Msg_Recovered) 'Green'
            }
            # 如果 Test-Path 失敗 (找不到 KeyToKey 執行檔)，什麼都不做
            # 因為在腳本開頭的「初始化檢查」已經警告過使用者了
        }
    }

} catch {
    # 捕捉所有未預期的錯誤
    Write-Host "`n[嚴重錯誤] 程式發生未預期的例外狀況：" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host "按 Enter 鍵離開..." -NoNewline
    Read-Host
}