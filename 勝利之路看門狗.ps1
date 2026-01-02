<#
    【勝利之路 & KeyToKey 看門狗 v1.0】
    Release Note:
    1. 正式發布 v1.0 版本。
    2. 整合所有穩定性修復與環境檢查功能。
    3. 新增「優雅退場」機制 (按 Q 鍵)。
#>

# ==========================================
# 0. Global Setup（全域環境設定)
# ==========================================
$ErrorActionPreference = 'Stop' 
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# 清除殘留變數
Remove-Variable SessionLog -Scope Global -ErrorAction SilentlyContinue
Remove-Variable LastReportLogIndex -Scope Global -ErrorAction SilentlyContinue
Remove-Variable LastBitmapCache -Scope Global -ErrorAction SilentlyContinue

try {
    # --- 特殊符號定義 ---
    $CR = [char]13; $LF = [char]10
    $Icon_Warn = [string][char]0x26A0 + [char]0xFE0F # ⚠️
    $Icon_Check  = [char]0x2705 # ✅
    $Icon_Cross  = [char]0x274C # ❌
    $Icon_Heart  = [string][char]0xD83E + [char]0xDEC0 # 🫀
    $Icon_Bullet = [char]0x2022 # •
    $Icon_Start  = [string][char]0x26A1 + [char]0x26BD # ⚡⚽
    $Icon_Stop   = [char]0x23F9 # ⏹️

    # --- 中文訊息設定 ---
    $Msg_Title_Start    = '看門狗 v1.0 已啟動'
    $Msg_Reason_Start   = '啟動通知'
    $Msg_Title_Stop     = '監控已手動停止'
    $Msg_Reason_Stop    = '使用者主動結束監控'
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
    $Msg_Stop_Monitor   = '監控已停止。按 Enter 鍵離開視窗...'
    $Msg_Status_OK      = '掛機運作中'
    $Msg_Sent_Report    = '已發送定期 Discord 報告'
    $Msg_KTK_Restart    = 'KeyToKey 重啟中...'
    $Msg_Wait_Load      = '等待 35 秒載入...'
    $Msg_Send_Key       = '發送按鍵'
    $Msg_Recovered      = '復原完畢'
    $Msg_Footer_Base    = 'Watchdog v1.0'
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

    # Webhook 讀取邏輯
    $ScriptPath = $MyInvocation.MyCommand.Path
    $ScriptDir  = Split-Path $ScriptPath -Parent
    $WebhookFile = Join-Path $ScriptDir 'webhook.txt'
    if (Test-Path $WebhookFile) {
        $DiscordWebhookUrl = (Get-Content $WebhookFile -Raw).Trim()
    } else { $DiscordWebhookUrl = '' }

    # 監控參數
    $LoopIntervalSeconds = 75  # 基礎間隔 75秒 + 處理時間 ≈ 90秒 (1:30)
    $FreezeThreshold = 3       
    $NoResponseThreshold = 3   
    $FreezeSimilarity = 98.5   

    # 初始化全域變數
    $Global:SessionLog = @()
    $Global:LastReportLogIndex = 0 
    $Global:LastHeartbeatTime = Get-Date
    $Global:HeartbeatInterval = 10 
    $Global:LastBitmapCache = $null 

    # ==========================================
    # 2. System Core（系統核心與 Windows API）
    # ==========================================
    try {
        $PSWindow = (Get-Host).UI.RawUI
        $BufferSize = $PSWindow.BufferSize; $BufferSize.Width = 120; $PSWindow.BufferSize = $BufferSize
        $WindowSize = $PSWindow.WindowSize; $WindowSize.Width = 120; $PSWindow.WindowSize = $WindowSize
    } catch {}

    if (!(Test-Path $ScreenshotDir)) { New-Item -ItemType Directory -Path $ScreenshotDir | Out-Null }
    Add-Type -AssemblyName System.Drawing, System.Net.Http, System.Windows.Forms

    # C# 核心代碼
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
        
        // 輸入法控制 API
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
    # 檢查是否已載入，避免重複執行時報錯
    if (-not ("Win32Tools" -as [type])) {
        Add-Type -TypeDefinition $Win32Code
    }

    try { [Console]::CursorVisible = $false } catch {}
    try { [Win32Tools]::SetProcessDPIAware() | Out-Null } catch {}

    # ==========================================
    # 3. Helpers（輔助功能函式庫）
    # ==========================================

    function Write-Log {
        param($Message, $Color='White', $ForceNewLine=$false)
        $Time = Get-Date -Format 'HH:mm:ss'
        $LogLine = '[' + $Time + '] ' + $Message
        if ($ForceNewLine) { Write-Host '' } 
        Write-Host ($LogLine + '          ') -ForegroundColor $Color
        $Global:SessionLog += $LogLine
        
        # 記憶體保護：限制日誌緩存最多 1000 行
        if ($Global:SessionLog.Count -gt 1000) { 
            $Global:SessionLog = $Global:SessionLog[-1000..-1] 
            $Global:LastReportLogIndex = [Math]::Max(0, $Global:LastReportLogIndex - ($Global:SessionLog.Count - 1000))
        }
    }

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

    # 強制鎖定英文輸入法
    function Ensure-English-IME {
        try {
            $HKL = [Win32Tools]::LoadKeyboardLayout("00000409", 1) 
            [Win32Tools]::ActivateKeyboardLayout($HKL, 0) | Out-Null
        } catch {}
    }

    function Send-Key-Native ($KeyName) {
        try {
            $KeyStr = '{' + $KeyName + '}'
            [System.Windows.Forms.SendKeys]::SendWait($KeyStr)
            return $true
        } catch { return $false }
    }

    function Show-Crash-Warning-GUI {
        param([string]$Reason)
        
        $Sym_Warn = [char]0x26A0; $Sym_Cancel = [char]0x2716
        $Color_Bg = [System.Drawing.Color]::FromArgb(30, 30, 30)
        $Color_Accent = [System.Drawing.Color]::FromArgb(255, 60, 60)
        $Color_TextPri = [System.Drawing.Color]::White
        $Color_TextSec = [System.Drawing.Color]::FromArgb(200, 200, 200)
        
        $Form = New-Object System.Windows.Forms.Form
        $Form.Size = New-Object System.Drawing.Size(600, 380)
        $Form.StartPosition = 'CenterScreen'; $Form.TopMost = $true; $Form.FormBorderStyle = 'None'
        $Form.BackColor = $Color_Accent; $Form.Padding = New-Object System.Windows.Forms.Padding(4)

        $MainPanel = New-Object System.Windows.Forms.Panel
        $MainPanel.Dock = 'Fill'; $MainPanel.BackColor = $Color_Bg; $Form.Controls.Add($MainPanel)

        $LblTitle = New-Object System.Windows.Forms.Label
        $LblTitle.Text = "$Sym_Warn 偵測到嚴重錯誤"
        $LblTitle.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 20, [System.Drawing.FontStyle]::Bold)
        $LblTitle.ForeColor = $Color_Accent; $LblTitle.AutoSize = $false
        $LblTitle.Size = New-Object System.Drawing.Size(592, 50); $LblTitle.Location = New-Object System.Drawing.Point(0, 30)
        $LblTitle.TextAlign = 'MiddleCenter'; $MainPanel.Controls.Add($LblTitle)

        $LblReason = New-Object System.Windows.Forms.Label
        $DispReason = if ($Reason.Length -gt 45) { $Reason.Substring(0, 42) + "..." } else { $Reason }
        $LblReason.Text = "$DispReason"
        $LblReason.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 12)
        $LblReason.ForeColor = $Color_TextSec; $LblReason.AutoSize = $false
        $LblReason.Size = New-Object System.Drawing.Size(592, 30); $LblReason.Location = New-Object System.Drawing.Point(0, 80)
        $LblReason.TextAlign = 'MiddleCenter'; $MainPanel.Controls.Add($LblReason)

        $LblCount = New-Object System.Windows.Forms.Label
        $LblCount.Text = "60"
        $LblCount.Font = New-Object System.Drawing.Font("Arial", 55, [System.Drawing.FontStyle]::Bold)
        $LblCount.ForeColor = $Color_TextPri; $LblCount.AutoSize = $false
        $LblCount.Size = New-Object System.Drawing.Size(592, 100); $LblCount.Location = New-Object System.Drawing.Point(0, 115)
        $LblCount.TextAlign = 'MiddleCenter'; $MainPanel.Controls.Add($LblCount)
        
        $LblSub = New-Object System.Windows.Forms.Label
        $LblSub.Text = "秒後將執行系統保護關機..."
        $LblSub.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 11)
        $LblSub.ForeColor = $Color_Accent; $LblSub.AutoSize = $false
        $LblSub.Size = New-Object System.Drawing.Size(592, 30); $LblSub.Location = New-Object System.Drawing.Point(0, 215)
        $LblSub.TextAlign = 'TopCenter'; $MainPanel.Controls.Add($LblSub)

        $BtnCancel = New-Object System.Windows.Forms.Button
        $BtnCancel.Text = "$Sym_Cancel 取消關機"
        $BtnCancel.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 16, [System.Drawing.FontStyle]::Bold)
        $BtnCancel.Size = New-Object System.Drawing.Size(260, 60)
        $BtnX = [int]((600 - 260) / 2)
        $BtnCancel.Location = New-Object System.Drawing.Point([int]($BtnX - 4), 270)
        $BtnCancel.BackColor = [System.Drawing.Color]::White; $BtnCancel.ForeColor = [System.Drawing.Color]::Black
        $BtnCancel.FlatStyle = 'Flat'; $BtnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $MainPanel.Controls.Add($BtnCancel)

        $Timer = New-Object System.Windows.Forms.Timer; $Timer.Interval = 1000; $Script:CountDown = 60
        $Timer.Add_Tick({
            $Script:CountDown--
            $LblCount.Text = "$Script:CountDown"
            if ($Script:CountDown -le 0) { $Timer.Stop(); $Form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $Form.Close() }
        })
        $Timer.Start()
        $Form.Add_Shown({ $BtnCancel.Focus() })
        $Result = $Form.ShowDialog(); $Timer.Stop(); $Form.Dispose(); return $Result
    }

    function Send-Discord-Report {
        param([string]$Title, [string]$Reason, [string]$ColorType='Green', [string[]]$ImagePaths=@(), [bool]$IsHeartbeat=$false)
        if ([string]::IsNullOrWhiteSpace($DiscordWebhookUrl) -or $DiscordWebhookUrl -eq 'YOUR_WEBHOOK_HERE') { return }
        
        if (!$IsHeartbeat) { Write-Log 'Uploading Report...' 'Cyan' $false }

        $LogPreviewLines = @()
        if ($IsHeartbeat) {
            $NewCount = $Global:SessionLog.Count
            if ($NewCount -gt $Global:LastReportLogIndex) {
                for ($k = $Global:LastReportLogIndex; $k -lt $NewCount; $k++) { $LogPreviewLines += $Global:SessionLog[$k] }
            }
            $Global:LastReportLogIndex = $NewCount
        } else {
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

        $ColorMap = @{ 'Green'=5763719; 'Red'=15548997; 'Yellow'=16705372; 'Blue'=5793266; 'Grey'=9807270 }
        
        $Duration = New-TimeSpan -Start $ScriptStartTime -End (Get-Date)
        $RunTimeStr = "{0:D2}小時{1:D2}分鐘" -f [int][Math]::Floor($Duration.TotalHours), $Duration.Minutes

        $DescHeader = ''
        $MentionContent = ''

        if ($IsHeartbeat) {
            $DescHeader = '**' + $Icon_Check + ' ' + $Msg_Discord_HBTxt + '**' + $LF + $Msg_Discord_SysOK + $LF + 
                          "(每 $Global:HeartbeatInterval 分鐘回報一次)" + $LF + $LF + '⏱️ **已運行時間**' + $LF + $RunTimeStr
        } else {
            $DescHeader = "**異常原因：**$LF" + $Reason + $LF + $LF + "⏳ **已掛機：**$LF" + $RunTimeStr
            if ($ColorType -eq 'Red' -or $ColorType -eq 'Yellow') { $MentionContent = "<@$DiscordUserID>" }
        }

        if ($ColorType -ne 'Yellow') {
            $EmbedDesc = $DescHeader + $LF + $LF + '**📋 ' + $Msg_Discord_Log + '**' + $LF + '```' + $LF + $LogPreview + $LF + '```'
        } else { $EmbedDesc = $DescHeader }

        $FooterTxt = $Msg_Footer_Base + ' ' + $Icon_Bullet + ' ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        $Embed = @{ title = $Title; description = $EmbedDesc; color = $ColorMap[$ColorType]; footer = @{ text = $FooterTxt } }
        $Payload = @{ content = $MentionContent; embeds = @($Embed) }
        $JsonPayload = $Payload | ConvertTo-Json -Depth 10 -Compress

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
            if (!$IsHeartbeat -and $ColorType -ne 'Yellow' -and (Test-Path $LogSavePath)) {
                $FS2 = [System.IO.File]::OpenRead($LogSavePath); $Streams += $FS2
                $TxtContent = New-Object System.Net.Http.StreamContent($FS2)
                $TxtContent.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse('text/plain')
                $Form.Add($TxtContent, "file_log", 'Watchdog_Log.txt')
            }
            $HttpClient.PostAsync($DiscordWebhookUrl, $Form).Result | Out-Null
        } catch { Write-Log "Discord 上傳失敗: $_" 'Red' $true } 
        finally {
            foreach ($s in $Streams) { $s.Close(); $s.Dispose() }
            if ($HttpClient) { $HttpClient.Dispose() }
            if ($Form) { $Form.Dispose() }
        }

        Start-Sleep -Seconds 1
        foreach ($Path in $ImagePaths) { if (Test-Path $Path) { try { Remove-Item $Path -Force -ErrorAction SilentlyContinue } catch {} } }
        if (Test-Path $LogSavePath) { try { Remove-Item $LogSavePath -Force -ErrorAction SilentlyContinue } catch {} }
    }

    function Suppress-Steam-Window {
        $Targets = [SteamBuster]::FindSteamWindows()
        $ClosedAny = $false
        foreach ($win in $Targets) {
            $Msg = "偵測到干擾視窗！標題: [$($win.Title)] (程式: $($win.ProcessName))"
            Write-Log ($Icon_Warn + ' ' + $Msg) 'Yellow' $true
            Send-Discord-Report -Title ($Icon_Warn + ' 異常徵兆警告') -Reason "$Msg`n(已執行自動關閉)" -ColorType 'Yellow'
            [SteamBuster]::CloseWindow($win.Handle)
            $ClosedAny = $true
            Write-Log ($Icon_Check + ' 已關閉視窗。') 'Green'
        }
        if ($ClosedAny) { Ensure-Game-TopMost }
    }

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

    function Get-PixelsFromBitmap ($Bitmap) {
        try {
            if (!$Bitmap) { return $null }
            $Small = $Bitmap.GetThumbnailImage(12, 12, $null, [IntPtr]::Zero)
            $Pixels = New-Object 'int[,]' 12, 12
            for ($x=0; $x -lt 12; $x++) { for ($y=0; $y -lt 12; $y++) { $Pixels[$x, $y] = $Small.GetPixel($x, $y).ToArgb() } }
            $Small.Dispose(); return ,$Pixels 
        } catch { return $null }
    }

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
    Write-Host '   Victory Road & KeyToKey Watchdog v1.0' -ForegroundColor Cyan
    Write-Host '   (Release Version)' -ForegroundColor Cyan
    Write-Host '==========================================' -ForegroundColor Cyan

    # 路徑防呆檢查
    if (!(Test-Path $KeyToKeyPath)) {
        Write-Host "⚠️ 警告：找不到 KeyToKey 執行檔！" -ForegroundColor Red
        Write-Host "路徑：$KeyToKeyPath" -ForegroundColor Red
        Write-Host "自動重啟功能將失效。" -ForegroundColor Yellow
        Write-Host "請修改腳本中的 `$KeyToKeyPath 變數。" -ForegroundColor Yellow
        Start-Sleep 3
    }

    if ([string]::IsNullOrWhiteSpace($DiscordWebhookUrl)) {
        Write-Host ''; Write-Host $Msg_Ask_Webhook -ForegroundColor Yellow
        $InputUrl = Read-Host 'URL'
        if (![string]::IsNullOrWhiteSpace($InputUrl)) {
            $DiscordWebhookUrl = $InputUrl.Trim()
            $DiscordWebhookUrl | Out-File -FilePath $WebhookFile -Encoding UTF8
            Write-Host $Msg_Webhook_Saved -ForegroundColor Green
        }
    }

    Write-Host ''; Write-Host '[設定] 當遊戲崩潰時，是否要執行電腦關機保護？ (按 Y 啟用，按其他鍵停用)' -ForegroundColor Yellow
    $ShutdownInput = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    $EnableShutdown = ($ShutdownInput.Character -eq 'y' -or $ShutdownInput.Character -eq 'Y')
    if ($EnableShutdown) { Write-Host 'Y (已啟用關機保護)' -ForegroundColor Red } else { Write-Host 'N (僅關閉程式)' -ForegroundColor Green }

    Write-Host ''; Write-Host '[設定] 請輸入 KTK 啟動熱鍵  [預設：F7]' -ForegroundColor Yellow
    $InputKey = Read-Host '請輸入'
    $TargetKeyName = if ([string]::IsNullOrWhiteSpace($InputKey)) { 'F7' } else { $InputKey.Trim().ToUpper() }
    Write-Host ('已設定按鍵: ' + $TargetKeyName) -ForegroundColor Green

    Write-Host ''; Write-Host '[設定] 請輸入 Discord 定期回報間隔 (分鐘) [預設：10 分鐘]' -ForegroundColor Yellow
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
    $Global:LastBitmapCache = Capture-ScreenBitmap
    $LastPixelData = Get-PixelsFromBitmap $Global:LastBitmapCache

    Write-Host ''; Write-Host '=== 監控開始 (按 Q 鍵停止並回報) ===' -ForegroundColor Cyan
    $Global:SessionLog = @()
    Send-Discord-Report -Title ($Icon_Start + ' ' + $Msg_Title_Start) -Reason $Msg_Reason_Start -ColorType 'Blue' -IsHeartbeat $true

    # 啟動前環境檢查 (不等待倒數)
    Ensure-Game-TopMost
    Ensure-English-IME
    if ((Get-Process -Name 'nie' -ErrorAction SilentlyContinue) -and !(Get-Process -Name 'KeyToKey' -ErrorAction SilentlyContinue)) {
        Write-Log "➤ 初始檢查：KeyToKey 未執行，嘗試啟動..." 'Yellow'
        if (Test-Path $KeyToKeyPath) { Start-Process $KeyToKeyPath; Start-Sleep 5 }
    }

    while ($true) {
        Ensure-Game-TopMost
        Ensure-English-IME
        
        # --- 倒數計時 ---
        for ($i = $LoopIntervalSeconds; $i -gt 0; $i--) {
            Suppress-Steam-Window
            
            # 優雅退場偵測 (按 Q 鍵)
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

            if ($LoopIntervalSeconds -gt 1) { $Percent = ($LoopIntervalSeconds - $i) / ($LoopIntervalSeconds - 1) } else { $Percent = 1 }
            $ProgressCount = [int][Math]::Floor($Percent * 20)
            
            $CheckGame = Get-Process -Name 'nie' -ErrorAction SilentlyContinue
            $CheckKTK = Get-Process -Name 'KeyToKey' -ErrorAction SilentlyContinue
            $StatusStr = ''
            if ($CheckGame) { if ($CheckGame.Responding) { $StatusStr += $Msg_Game_Run } else { $StatusStr += $Msg_Game_NoResp + ' ' + $Icon_Warn } } else { $StatusStr += $Msg_Game_Lost + ' ' + $Icon_Cross }
            $StatusStr += ' | '
            if ($CheckKTK)  { $StatusStr += $Msg_KTK_Run } else { $StatusStr += $Msg_KTK_Err + ' ' + $Icon_Warn }

            for ($blink = 0; $blink -lt 2; $blink++) {
                $BarStr = ''
                if ($ProgressCount -ge 20) { $BarStr = '=' * 20 } 
                else { if ($ProgressCount -gt 0) { if ($blink -eq 0) { $BarStr = '=' * $ProgressCount } else { $BarStr = '=' * ($ProgressCount - 1) + ' ' } } }
                $Bar = '[' + $BarStr + (' ' * (20 - $BarStr.Length)) + ']'
                Write-Host ($CR + $Bar + " 倒數 $i 秒... (按 Q 停止) [ $StatusStr ]       ") -NoNewline -ForegroundColor Gray
                Start-Sleep -Milliseconds 500
            }
        }
        Write-Host '' 

        # --- 核心檢測邏輯 ---
        $GameProcess = Get-Process -Name 'nie' -ErrorAction SilentlyContinue
        $KTKProcess = Get-Process -Name 'KeyToKey' -ErrorAction SilentlyContinue
        $ErrorTriggered = $false; $ErrorReason = ''
        
        $CurrentBitmap = Capture-ScreenBitmap
        $CurrentPixelData = Get-PixelsFromBitmap $CurrentBitmap
        $ReportImages = @() 

        # 1. 檢測：消失 + 系統日誌
        if (!$GameProcess) { 
            $ErrorTriggered = $true
            $TimeLimit = (Get-Date).AddMinutes(-5) 
            $KernelErrors = Get-WinEvent -FilterHashtable @{LogName='System'; Id=141,4101,41,117,10016,1001} -ErrorAction SilentlyContinue | Where-Object { $_.TimeCreated -gt $TimeLimit }
            
            if ($KernelErrors) {
                $RecentError = $KernelErrors | Select-Object -First 1
                $SysErrMsg = $Msg_Err_Reason + ' ' + $RecentError.Id + ')'
                $ErrorReason = $SysErrMsg
                Write-Log ($Icon_Cross + ' 偵測到程式消失，且發現系統錯誤 ID: ' + $RecentError.Id) 'Red'
            } else { $ErrorReason = $Msg_Err_Crash }
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
                if ($Global:LastBitmapCache) { $PathPrev = Save-BitmapToFile $Global:LastBitmapCache 'Freeze_Prev'; if ($PathPrev) { $ReportImages += $PathPrev } }
                if ($CurrentBitmap) { $PathCurr = Save-BitmapToFile $CurrentBitmap 'Freeze_Curr'; if ($PathCurr) { $ReportImages += $PathCurr } }

                if ($FreezeCount -ge $FreezeThreshold) { $ErrorTriggered = $true; $ErrorReason = $Msg_Err_Freeze } 
                else { Send-Discord-Report -Title ($Icon_Warn + ' ' + $Msg_Discord_Title_W) -Reason "畫面相似度過高 ($Similarity%) - 累積 $FreezeCount/$FreezeThreshold" -ColorType 'Yellow' -ImagePaths $ReportImages }
            } else { $FreezeCount = 0 }
        }

        # 4. 檢測：系統錯誤 (含凍結時)
        $TimeLimit = (Get-Date).AddMinutes(-5)
        $KernelErrors = Get-WinEvent -FilterHashtable @{LogName='System'; Id=141,4101,41,117,10016,1001} -ErrorAction SilentlyContinue | Where-Object { $_.TimeCreated -gt $TimeLimit }
        
        if ($KernelErrors) {
            $RecentError = $KernelErrors | Select-Object -First 1
            $SysErrMsg = $Msg_Err_Reason + ' ' + $RecentError.Id + ')'
            if ($ErrorTriggered) {
                if ($ErrorReason -notmatch $RecentError.Id) {
                    $ErrorReason += "`n[系統紀錄] $SysErrMsg"
                    Write-Log ($Icon_Cross + ' 補充偵測：' + $SysErrMsg) 'Red'
                }
            } else {
                $ErrorTriggered = $true; $ErrorReason = $SysErrMsg
                Write-Log ($Icon_Cross + ' ' + $Msg_Err_Sys + ' ' + $RecentError.Id) 'Red'
            }
        }
        
        # --- 異常處理流程 ---
        if ($ErrorTriggered) {
            $FinalDur = New-TimeSpan -Start $ScriptStartTime -End (Get-Date)
            $FinalTimeStr = "{0:D2}小時{1:D2}分鐘" -f [int][Math]::Floor($FinalDur.TotalHours), $FinalDur.Minutes

            Write-Log ($Icon_Cross + ' ' + $Msg_Prot_Trig + ' ' + $ErrorReason) 'Red' $true
            if ($EnableShutdown) { Write-Log "➤ 將執行自動關機程序" 'Yellow' }
            Write-Log "⏱️ 本次共掛機：$FinalTimeStr" 'Cyan'

            if ($ReportImages.Count -eq 0 -and $CurrentBitmap) {
                $PathCrash = Save-BitmapToFile $CurrentBitmap 'Crash'
                if ($PathCrash) { $ReportImages += $PathCrash }
            }
            # KTK 關閉保護
            if ($KTKProcess) { try { Stop-Process -Name 'KeyToKey' -Force -ErrorAction Stop } catch { Write-Log "⚠️ 無法強制關閉 KeyToKey: $($_.Exception.Message)" 'Yellow' } }
            Stop-Process -Name 'nie' -Force -ErrorAction SilentlyContinue
            
            $DiscordReason = if ($EnableShutdown) { "$ErrorReason`n(已執行自動關機程序)" } else { $ErrorReason }
            Send-Discord-Report -Title ($Icon_Cross + ' ' + $Msg_Discord_Title) -Reason $DiscordReason -ColorType 'Red' -ImagePaths $ReportImages
            
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

        if ($Global:LastBitmapCache) { $Global:LastBitmapCache.Dispose() }
        $Global:LastBitmapCache = $CurrentBitmap 
        $LastPixelData = $CurrentPixelData
        if (!$ErrorTriggered -and $KTKProcess) { Write-Log ('➤ ' + $Msg_Status_OK) 'DarkGray' }

        $TimeSinceLastHeartbeat = (Get-Date) - $Global:LastHeartbeatTime
        if ($TimeSinceLastHeartbeat.TotalMinutes -ge $Global:HeartbeatInterval) {
            $HbDur = New-TimeSpan -Start $ScriptStartTime -End (Get-Date)
            $HbTimeStr = "{0:D2}小時{1:D2}分鐘" -f [int][Math]::Floor($HbDur.TotalHours), $HbDur.Minutes
            Write-Log ('➤ ' + $Msg_Sent_Report + " (已運行時間：$HbTimeStr)") 'Cyan'
            $HbPath = Save-BitmapToFile $Global:LastBitmapCache 'Heartbeat'
            $HbPaths = if ($HbPath) { @($HbPath) } else { @() }
            Send-Discord-Report -Title ($Icon_Heart + ' ' + $Msg_Discord_HB) -Reason 'Heartbeat' -ColorType 'Green' -ImagePaths $HbPaths -IsHeartbeat $true
            $Global:LastHeartbeatTime = Get-Date
        }

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