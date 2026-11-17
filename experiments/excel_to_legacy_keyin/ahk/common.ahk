; =============================================================================
; common.ahk - AutoHotkey v2 공통 함수 (각 예제에서 #Include)
;   설정 / 실행·연결·로그인 / 컨트롤 찾기 / 주문 입력·저장 / 메시지박스 / CSV
; =============================================================================
#Requires AutoHotkey v2.0
SetTitleMatchMode 2          ; 2 = 제목(또는 컨트롤 텍스트)에 '포함'되면 일치
SetControlDelay 20           ; Control* 함수 사이 지연(ms). 느린 프로그램이면 늘리기

; -----------------------------------------------------------------[설정]
MAIN_TITLE  := "주문 등록 - Legacy Order System"
LOGIN_TITLE := "로그인 - Legacy Order System"
LOGIN_ID := "admin", LOGIN_PW := "1234"
DATA_DIR := A_ScriptDir "\..\data"
APP_CMD := 'powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "'
         . A_ScriptDir '\..\mock_app\LegacyOrderApp.ps1"'
; Edit 컨트롤을 위치순(위→아래, 왼→오른쪽)으로 정렬했을 때의 이름
EDIT_ORDER := ["고객코드", "고객명", "전화번호", "주문일자", "수량", "단가", "금액", "비고"]
; 사용자가 입력하는 Edit (금액은 프로그램이 계산하는 읽기 전용)
INPUT_FIELDS := ["고객코드", "고객명", "전화번호", "주문일자", "수량", "단가", "비고"]
; 지역 콤보 항목 (사전 검증, 키보드 입력의 ↓ 횟수 계산용)
REGIONS := ["서울", "부산", "대구", "인천", "광주", "대전", "울산"]
; -----------------------------------------------------------------------

; ---------------------------------------------------------- 실행 / 연결
; 프로그램이 떠 있으면 연결, 없으면 실행 → 로그인 → 메인 창 HWND 반환
StartOrConnect() {
    if !WinExist(MAIN_TITLE) && !WinExist(LOGIN_TITLE)
        Run APP_CMD, , "Hide"

    deadline := A_TickCount + 20000
    while !WinExist(MAIN_TITLE) {
        if (A_TickCount > deadline)
            throw Error("메인 창이 나타나지 않았습니다.")
        if (h := WinExist(LOGIN_TITLE))
            DoLogin(h)
        Sleep 200
    }
    main := WinExist(MAIN_TITLE)
    CloseAllMsgBoxes(main)                  ; 이전 실행에서 남은 메시지박스 정리
    return main
}

DoLogin(loginHwnd) {
    edits := GetControlsSorted(loginHwnd, "i)edit|textbox")   ; [아이디, 비밀번호]
    ControlSetText LOGIN_ID, edits[1].hwnd
    ControlSetText LOGIN_PW, edits[2].hwnd
    ControlClick "확인", loginHwnd                             ; 버튼 텍스트로 찾기 ("확인(&O)" 포함)
    WinWaitClose loginHwnd, , 10
}

; ---------------------------------------------------------- 컨트롤 찾기
; 클래스명이 classPattern(정규식)에 맞는 컨트롤을 '화면 위치 순서'로 반환
;   VB6 Label / Delphi TLabel 은 핸들이 없어 라벨로 찾을 수 없음 → 위치 순서가 가장 범용적
;   VB6: ThunderRT6TextBox, Delphi: TEdit/TMemo, WinForms: WindowsForms10.EDIT... 모두 매칭
GetControlsSorted(winHwnd, classPattern) {
    items := []
    for ctrl in WinGetControlsHwnd(winHwnd) {
        if !RegExMatch(WinGetClass(ctrl), classPattern) || !ControlGetVisible(ctrl)
            continue
        ControlGetPos &x, &y, , , ctrl
        items.Push({hwnd: ctrl, x: x, y: y})
    }
    return SortByPosition(items)
}

; {x, y} 배열을 위치순으로 정렬 (삽입 정렬)
SortByPosition(items) {
    loop items.Length - 1 {
        i := A_Index + 1, cur := items[i], j := i - 1
        while (j >= 1 && IsBefore(cur, items[j])) {
            items[j + 1] := items[j], j--
        }
        items[j + 1] := cur
    }
    return items
}
; y 가 8px 넘게 차이 나면 위쪽 먼저, 같은 줄이면 왼쪽 먼저
IsBefore(a, b) => Abs(a.y - b.y) > 8 ? a.y < b.y : a.x < b.x

; 메인 창의 Edit/콤보를 Map(이름 → HWND) 로 반환. 버튼·라디오는 텍스트로 찾으므로 제외
GetFields(main) {
    edits := GetControlsSorted(main, "i)edit|textbox|memo")
    if (edits.Length != EDIT_ORDER.Length)
        throw Error("Edit 개수가 예상과 다릅니다: " edits.Length)
    f := Map()
    for i, name in EDIT_ORDER
        f[name] := edits[i].hwnd
    f["지역"] := GetControlsSorted(main, "i)combo")[1].hwnd
    return f
}

; ---------------------------------------------------------- 주문 입력 / 저장
; 프로그램에 넣기 전에 걸러낼 수 있는 문제 → 사유 문자열, 문제없으면 ""
CheckOrder(rec) {
    if (rec["고객코드"] = "")
        return "고객코드 없음"
    if !IsInteger(rec["수량"])
        return "수량이 숫자가 아님(" rec["수량"] ")"
    if !IndexOf(REGIONS, rec["지역"])
        return "지역 목록에 없음(" rec["지역"] ")"
    return ""
}

; 화면 초기화 후 rec(Map: 필드명 → 값)을 컨트롤에 직접 입력. 입력값은 다시 읽어 검증
FillOrder(main, rec) {
    f := GetFields(main)
    ControlClick "초기화", main
    Sleep 150
    for name in INPUT_FIELDS {
        ControlSetText rec[name], f[name]
        if (ControlGetText(f[name]) != rec[name])
            throw Error(name " 입력 확인 실패")
    }
    ControlChooseString rec["지역"], f["지역"]                 ; 목록에 없으면 예외
    ControlSetChecked 1, (rec["고객구분"] = "법인" ? "법인" : "개인"), main
    ControlSetChecked rec["VIP"] = "Y", "VIP", main
    return f
}

; 저장 버튼 → 메시지박스로 판정 → "성공" 또는 "실패: 사유"
SaveOrder(main) {
    ControlClick "저장", main
    box := WaitMsgBox(main, 20)
    if !box
        throw Error("저장 응답(메시지박스) 없음")
    CloseMsgBox(box)
    return (box.title = "확인" && InStr(box.text, "저장")) ? "성공" : "실패: " box.text
}

; ---------------------------------------------------------- 메시지박스
; 앱이 띄운 메시지박스(#32770) 대기 → {hwnd, title, text} 또는 "" (시간초과)
WaitMsgBox(main, timeoutSec := 15) {
    h := WinWait("ahk_class #32770 ahk_pid " WinGetPID(main), , timeoutSec)
    if !h
        return ""
    text := ""
    for ctrl in WinGetControlsHwnd(h)
        if InStr(WinGetClass(ctrl), "Static") && (t := ControlGetText(ctrl)) != ""
            text .= t " "
    return {hwnd: h, title: WinGetTitle(h), text: Trim(text)}
}

; 첫 번째 버튼(확인/OK) 클릭. "예" 를 누르려면 button := "예"
CloseMsgBox(box, button := "Button1") {
    ControlClick button, box.hwnd
    WinWaitClose box.hwnd, , 5
}

CloseAllMsgBoxes(main) {
    while (box := WaitMsgBox(main, 0.5))
        CloseMsgBox(box)
}

; ---------------------------------------------------------- 유틸
; UTF-8 CSV → Map 배열 (첫 줄 = 헤더).  ※ 셀 안의 줄바꿈은 지원하지 않는 단순 버전
ReadCsv(path) {
    rows := [], headers := []
    loop parse FileRead(path, "UTF-8"), "`n", "`r" {
        if (Trim(A_LoopField) = "")
            continue
        fields := []
        loop parse A_LoopField, "CSV"
            fields.Push(Trim(A_LoopField))
        if !headers.Length {
            headers := fields
            continue
        }
        m := Map()
        for i, h in headers
            m[h] := fields.Has(i) ? fields[i] : ""
        rows.Push(m)
    }
    return rows
}

; 배열에서 값의 위치(1부터), 없으면 0
IndexOf(arr, value) {
    for i, v in arr
        if (v = value)
            return i
    return 0
}

Timestamp() => FormatTime(, "yyyy-MM-dd HH:mm:ss")
