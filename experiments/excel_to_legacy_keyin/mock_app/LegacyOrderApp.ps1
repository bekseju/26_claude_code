<#
  LegacyOrderApp.ps1  -  "옛날 VB/Delphi 프로그램"을 흉내 내는 연습용 입력 프로그램

  * 실제 VB6 / Delphi 프로그램처럼 Win32 네이티브 컨트롤(Edit, ComboBox, Button,
    #32770 메시지박스, Win32 메뉴)을 사용합니다.  -> pywinauto win32 backend / AutoHotkey 로 제어 가능
  * 로그인 창 -> 주문 등록 창 -> 저장 시 "처리중" 지연 -> 성공/오류 메시지박스
  * 저장된 데이터는 saved_records.csv 에 누적되어 자동화 결과를 검증할 수 있습니다.

  실행:  run_mock_app.bat   (또는)  powershell -STA -ExecutionPolicy Bypass -File LegacyOrderApp.ps1 [-NoLogin] [-NoDelay]
  로그인: admin / 1234

  ※ 자동화 스크립트(python/, ahk/)가 창 제목·컨트롤 텍스트·위치·Tab 순서·메시지 문구에 의존합니다.
     이 값들을 바꾸면 자동화 코드도 함께 고쳐야 합니다.
#>
param(
    [switch]$NoLogin,   # 로그인 창 생략
    [switch]$NoDelay    # 저장 시 "처리중" 지연 생략
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# 일부러 EnableVisualStyles() 를 호출하지 않음 -> 옛날 Win32 느낌

# ============================================================== 설정
$script:Config = @{
    CsvPath    = Join-Path $PSScriptRoot 'saved_records.csv'
    Regions    = @('서울', '부산', '대구', '인천', '광주', '대전', '울산')
    LoginId    = 'admin'
    LoginPw    = '1234'
    LoginTitle = '로그인 - Legacy Order System'
    MainTitle  = '주문 등록 - Legacy Order System v1.0'
}

# 주문 등록 화면 레이아웃: 첫 줄 y / 줄 간격 / 라벨·입력칸·도움말 x 좌표
$script:Layout = @{ Top = 15; RowHeight = 30; LabelX = 20; InputX = 110; HintX = 240 }

# ============================================================== 컨트롤 생성 도우미
function New-Point($x, $y) { New-Object System.Drawing.Point($x, $y) }
function New-Size($w, $h)  { New-Object System.Drawing.Size($w, $h) }

function Add-Label($parent, $text, $x, $y, $w = 85) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $text; $l.Location = New-Point $x ($y + 3); $l.Size = New-Size $w 18   # 입력칸과 글자 높이 맞춤
    $parent.Controls.Add($l); return $l
}

function Add-TextBox($parent, $x, $y, $w, $tabIndex) {
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location = New-Point $x $y; $t.Size = New-Size $w 22; $t.TabIndex = $tabIndex
    $parent.Controls.Add($t); return $t
}

function Add-Button($parent, $text, $x, $y, $tabIndex, $w = 90) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $text; $b.Location = New-Point $x $y; $b.Size = New-Size $w 28; $b.TabIndex = $tabIndex
    $b.FlatStyle = 'System'                         # FlatStyle=System -> 순수 Win32 BUTTON
    $parent.Controls.Add($b); return $b
}

# 라디오/체크박스 공통 (FlatStyle=System 이어야 BM_GETCHECK 등 Win32 메시지가 동작)
function Add-Choice($parent, $type, $text, $x, $y, $tabIndex) {
    $c = New-Object "System.Windows.Forms.$type"
    $c.Text = $text; $c.FlatStyle = 'System'
    $c.Location = New-Point $x $y; $c.Size = New-Size 70 22; $c.TabIndex = $tabIndex
    $parent.Controls.Add($c); return $c
}

function Add-DropDownList($parent, $items, $x, $y, $w, $tabIndex) {
    $c = New-Object System.Windows.Forms.ComboBox
    $c.DropDownStyle = 'DropDownList'; $c.FlatStyle = 'System'
    $c.Location = New-Point $x $y; $c.Size = New-Size $w 22; $c.TabIndex = $tabIndex
    [void]$c.Items.AddRange($items)
    $parent.Controls.Add($c); return $c
}

# n 번째 줄의 라벨을 만들고 그 줄의 y 좌표를 반환
function Add-Row($parent, $row, $labelText) {
    $y = $script:Layout.Top + $row * $script:Layout.RowHeight
    [void](Add-Label $parent $labelText $script:Layout.LabelX $y)
    return $y
}

function Show-Msg($owner, $text, $title, $icon = 'Information', $buttons = 'OK') {
    return [System.Windows.Forms.MessageBox]::Show($owner, $text, $title, $buttons, $icon)
}

# ============================================================== 데이터 (CSV)
function Read-SavedRecords {
    if (Test-Path $script:Config.CsvPath) { return @(Import-Csv -Path $script:Config.CsvPath -Encoding UTF8) }
    return @()
}

function Save-Record($record) {
    $record | Export-Csv -Path $script:Config.CsvPath -Encoding UTF8 -NoTypeInformation -Append
}

# ============================================================== 로그인 창
function Show-Login {
    $f = New-Object System.Windows.Forms.Form
    $f.Text = $script:Config.LoginTitle
    $f.ClientSize = New-Size 280 130
    $f.FormBorderStyle = 'FixedDialog'; $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $f.StartPosition = 'CenterScreen'

    [void](Add-Label $f '아이디' 20 20 60)
    $id = Add-TextBox $f 90 20 160 0
    [void](Add-Label $f '비밀번호' 20 52 60)
    $pw = Add-TextBox $f 90 52 160 1
    $pw.PasswordChar = '*'
    $ok = Add-Button $f '확인(&O)' 70 88 2 80
    $cancel = Add-Button $f '취소' 160 88 3 80
    $f.AcceptButton = $ok; $f.CancelButton = $cancel

    $ok.Add_Click({
        if ($id.Text -eq $script:Config.LoginId -and $pw.Text -eq $script:Config.LoginPw) {
            $f.DialogResult = 'OK'; $f.Close()
        } else {
            [void](Show-Msg $f '아이디 또는 비밀번호가 올바르지 않습니다.' '로그인 실패' 'Error')
            $pw.Text = ''; [void]$pw.Focus()
        }
    })
    $cancel.Add_Click({ $f.DialogResult = 'Cancel'; $f.Close() })
    return ($f.ShowDialog() -eq 'OK')
}

# ============================================================== 주문 등록 창 - 화면 구성
# MainMenu = 진짜 Win32 메뉴 -> pywinauto menu_select 가능
function New-MainMenu {
    $menu = New-Object System.Windows.Forms.MainMenu
    $file = $menu.MenuItems.Add('파일(&F)')
    $m = @{ Menu = $menu }
    $m.New = $file.MenuItems.Add('새로 입력(&N)'); $m.New.Shortcut = 'CtrlN'
    $m.Del = $file.MenuItems.Add('저장 데이터 모두 삭제(&D)')
    [void]$file.MenuItems.Add('-')
    $m.Exit = $file.MenuItems.Add('종료(&X)')
    $m.About = $menu.MenuItems.Add('도움말(&H)').MenuItems.Add('프로그램 정보(&A)')
    return $m
}

# 입력 필드 생성 (TabIndex 순서 = 키보드 Tab 이동 순서. 자동화 스크립트가 위치·순서에 의존)
function New-OrderFields($f) {
    $L = $script:Layout; $x = $L.InputX
    $ui = @{}

    $y = Add-Row $f 0 '고객코드'
    $ui.Code = Add-TextBox $f $x $y 120 0; $ui.Code.MaxLength = 5
    [void](Add-Label $f '(예: C0001)' $L.HintX $y 100)

    $y = Add-Row $f 1 '고객명';   $ui.Name   = Add-TextBox $f $x $y 200 1
    $y = Add-Row $f 2 '전화번호'; $ui.Phone  = Add-TextBox $f $x $y 200 2
    $y = Add-Row $f 3 '지역';     $ui.Region = Add-DropDownList $f $script:Config.Regions $x $y 120 3

    $y = Add-Row $f 4 '고객구분'
    $ui.Person = Add-Choice $f 'RadioButton' '개인(&P)' 110 $y 4; $ui.Person.Checked = $true
    $ui.Corp   = Add-Choice $f 'RadioButton' '법인(&B)' 185 $y 5
    $ui.Vip    = Add-Choice $f 'CheckBox'    'VIP(&V)'  265 $y 6

    $y = Add-Row $f 5 '주문일자'
    $ui.Date = Add-TextBox $f $x $y 120 7
    [void](Add-Label $f '(YYYY-MM-DD)' $L.HintX $y 100)

    $y = Add-Row $f 6 '수량'; $ui.Qty = Add-TextBox $f $x $y 80 8

    $y = Add-Row $f 7 '단가'; $ui.Price = Add-TextBox $f $x $y 100 9
    [void](Add-Label $f '금액' 225 $y 35)
    $ui.Amount = Add-TextBox $f 260 $y 100 99                 # 자동 계산 (읽기 전용, Tab 이동 제외)
    $ui.Amount.ReadOnly = $true; $ui.Amount.TabStop = $false; $ui.Amount.TextAlign = 'Right'

    $y = Add-Row $f 8 '비고'
    $ui.Memo = Add-TextBox $f $x $y 250 10; $ui.Memo.Multiline = $true; $ui.Memo.Height = 50

    $ui.Save  = Add-Button $f '저장(&S)'   20 330 11
    $ui.Clear = Add-Button $f '초기화(&C)' 120 330 12
    $ui.Exit  = Add-Button $f '종료(&X)'   270 330 13
    return $ui
}

function New-RecordList($f) {
    $lv = New-Object System.Windows.Forms.ListView
    $lv.View = 'Details'; $lv.FullRowSelect = $true; $lv.GridLines = $true; $lv.TabStop = $false
    $lv.Location = New-Point 380 15; $lv.Size = New-Size 365 345
    foreach ($c in @('고객코드', '고객명', '지역', '수량', '금액')) { [void]$lv.Columns.Add($c, 70) }
    $f.Controls.Add($lv)
    return $lv
}

# ============================================================== 주문 등록 창 - 동작
function Update-RecordList($ui, $records) {
    $ui.List.Items.Clear()
    foreach ($r in $records) {
        $it = New-Object System.Windows.Forms.ListViewItem($r.'고객코드')
        [void]$it.SubItems.AddRange([string[]]@($r.'고객명', $r.'지역', $r.'수량', $r.'금액'))
        [void]$ui.List.Items.Add($it)
    }
    $ui.Status.Text = "등록 건수: $($records.Count)"
}

function Clear-OrderForm($ui) {
    foreach ($k in 'Code', 'Name', 'Phone', 'Date', 'Qty', 'Price', 'Amount', 'Memo') { $ui[$k].Text = '' }
    $ui.Region.SelectedIndex = -1
    $ui.Person.Checked = $true; $ui.Vip.Checked = $false
    [void]$ui.Code.Focus()
}

# 수량 × 단가 -> 금액 (자동화 후 "읽어서 검증" 연습용)
function Update-Amount($ui) {
    $q = 0; $p = 0
    if ([int]::TryParse($ui.Qty.Text, [ref]$q) -and [int]::TryParse($ui.Price.Text, [ref]$p)) {
        $ui.Amount.Text = ($q * $p).ToString('N0')
    } else { $ui.Amount.Text = '' }
}

# 입력값 검증 (위에서부터 순서대로 검사)
#   오류: @{ Error = 메시지; Control = 포커스 줄 컨트롤 }   정상: @{ Code; Date; Qty; Price }
function Test-OrderInput($ui, $records) {
    $code = $ui.Code.Text.Trim().ToUpper()
    $q = 0; $p = 0; $d = [datetime]::MinValue
    $fail = { param($msg, $ctl) @{ Error = $msg; Control = $ctl } }

    if ($code -notmatch '^C\d{4}$') { return & $fail '고객코드 형식이 올바르지 않습니다. (예: C0001)' $ui.Code }
    if (@($records | Where-Object { $_.'고객코드' -eq $code }).Count -gt 0) { return & $fail "이미 등록된 고객코드입니다: $code" $ui.Code }
    if ($ui.Name.Text.Trim() -eq '') { return & $fail '고객명을 입력하세요.' $ui.Name }
    if ($ui.Region.SelectedIndex -lt 0) { return & $fail '지역을 선택하세요.' $ui.Region }
    if (-not [datetime]::TryParseExact($ui.Date.Text.Trim(), 'yyyy-MM-dd', $null, 'None', [ref]$d)) {
        return & $fail '주문일자 형식이 올바르지 않습니다. (YYYY-MM-DD)' $ui.Date
    }
    if (-not [int]::TryParse($ui.Qty.Text.Trim(), [ref]$q) -or $q -le 0) { return & $fail '수량은 1 이상의 숫자여야 합니다.' $ui.Qty }
    if (-not [int]::TryParse($ui.Price.Text.Trim(), [ref]$p) -or $p -lt 0) { return & $fail '단가는 숫자여야 합니다.' $ui.Price }
    return @{ Code = $code; Date = $d; Qty = $q; Price = $p }
}

function ConvertTo-OrderRecord($ui, $v) {
    [pscustomobject][ordered]@{
        '고객코드' = $v.Code
        '고객명'   = $ui.Name.Text.Trim()
        '전화번호' = $ui.Phone.Text.Trim()
        '지역'     = [string]$ui.Region.SelectedItem
        '고객구분' = $(if ($ui.Corp.Checked) { '법인' } else { '개인' })
        'VIP'      = $(if ($ui.Vip.Checked) { 'Y' } else { 'N' })
        '주문일자' = $v.Date.ToString('yyyy-MM-dd')
        '수량'     = $v.Qty
        '단가'     = $v.Price
        '금액'     = $v.Qty * $v.Price
        '비고'     = ($ui.Memo.Text -replace "`r?`n", ' / ')
        '등록시각' = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    }
}

function Invoke-SaveOrder($f, $ui, $records) {
    $v = Test-OrderInput $ui $records
    if ($v.Error) {
        [void](Show-Msg $f $v.Error '입력 오류' 'Warning')
        [void]$v.Control.Focus()
        return
    }

    # 옛날 프로그램처럼 "처리중..." 동안 화면이 멈춤 (UI 스레드 블로킹)
    $f.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $ui.Status.Text = '처리중...'; $f.Refresh()
    if (-not $NoDelay) { Start-Sleep -Milliseconds (Get-Random -Minimum 300 -Maximum 1500) }
    $f.Cursor = [System.Windows.Forms.Cursors]::Default

    $rec = ConvertTo-OrderRecord $ui $v
    [void]$records.Add($rec)
    Save-Record $rec
    Update-RecordList $ui $records
    [void](Show-Msg $f "저장되었습니다. ($($v.Code))" '확인')
    Clear-OrderForm $ui
}

function Clear-AllRecords($f, $ui, $records) {
    if ((Show-Msg $f '저장된 데이터를 모두 삭제할까요?' '삭제 확인' 'Question' 'YesNo') -ne 'Yes') { return }
    $records.Clear()
    if (Test-Path $script:Config.CsvPath) { Remove-Item $script:Config.CsvPath }
    Update-RecordList $ui $records
}

function Show-Main {
    $records = New-Object System.Collections.ArrayList
    foreach ($r in (Read-SavedRecords)) { [void]$records.Add($r) }

    $f = New-Object System.Windows.Forms.Form
    $f.Text = $script:Config.MainTitle
    $f.ClientSize = New-Size 760 420
    $f.FormBorderStyle = 'FixedSingle'; $f.MaximizeBox = $false
    $f.StartPosition = 'CenterScreen'
    $f.KeyPreview = $true

    $menu = New-MainMenu
    $f.Menu = $menu.Menu
    $ui = New-OrderFields $f
    $ui.List = New-RecordList $f
    $ui.Status = Add-Label $f '' 380 370 365
    Update-RecordList $ui $records

    # 이벤트 연결 (핸들러는 Show-Main 이 실행 중일 때 호출되므로 $f/$ui/$records 에 접근 가능)
    $ui.Qty.Add_TextChanged({ Update-Amount $ui })
    $ui.Price.Add_TextChanged({ Update-Amount $ui })
    # 고객코드는 포커스를 벗어날 때 대문자로 변환 (옛날 프로그램의 LostFocus 처리 흉내)
    $ui.Code.Add_Leave({ $ui.Code.Text = $ui.Code.Text.Trim().ToUpper() })

    $ui.Save.Add_Click({ Invoke-SaveOrder $f $ui $records })
    $ui.Clear.Add_Click({ Clear-OrderForm $ui })
    $ui.Exit.Add_Click({ $f.Close() })
    $menu.New.Add_Click({ Clear-OrderForm $ui })
    $menu.Del.Add_Click({ Clear-AllRecords $f $ui $records })
    $menu.Exit.Add_Click({ $f.Close() })
    $menu.About.Add_Click({ [void](Show-Msg $f "Legacy Order System v1.0`n자동화 연습용 가짜 레거시 프로그램" '프로그램 정보') })
    $f.Add_Shown({ [void]$ui.Code.Focus() })
    [void]$f.ShowDialog()
}

if ($NoLogin -or (Show-Login)) { Show-Main }
