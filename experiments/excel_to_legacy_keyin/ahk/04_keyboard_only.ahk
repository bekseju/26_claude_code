; =============================================================================
; [4단계] 키보드만으로 입력 - 컨트롤 인식이 안 될 때의 '최후의 수단'
;   (그리드/커스텀 컨트롤, Citrix/원격 화면, KeyPress 이벤트에서만 값을 받는 필드 등)
;
;   WinActivate + WinWaitActive : 반드시 창을 맨 앞으로
;   SendText "값"               : 문자 그대로 입력 (특수문자 해석 안 함)
;   클립보드 + ^v               : IME(한/영) 상태와 무관하게 가장 안전 → 한글에 사용
;   {Tab}                        : 다음 필드  ← Tab 순서를 미리 손으로 확인해 둘 것
;   콤보 {Home}{Down n}, 라디오 방향키, 체크박스 {Space}, 버튼 Alt+단축키(!s)
;
;   ※ 실제 키 입력이므로 실행 중에는 마우스/키보드를 건드리지 마세요.
; =============================================================================
#Requires AutoHotkey v2.0
#Include common.ahk
SetKeyDelay 30

rec := Map("고객코드", "C9102", "고객명", "AHK김키보드", "전화번호", "010-3333-0000", "지역", "대전"
         , "고객구분", "법인", "VIP", "Y", "주문일자", "2026-11-07", "수량", "3", "단가", "15000"
         , "비고", "AHK 키보드 전용 입력")
STEP := 150                                       ; 키 입력 사이 대기(ms)

; 키 입력 후 잠깐 대기, tab := true 면 다음 필드로 이동
Keys(keys, tab := true) {
    Send keys
    Sleep STEP
    if tab
        Send "{Tab}"
    Sleep STEP
}
TypeText(text) {                                  ; 영문/숫자: 직접 타이핑
    SendText text
    Keys("")
}
PasteText(text, tab := true) {                    ; 한글: 클립보드 붙여넣기 (기존 내용 덮어씀)
    A_Clipboard := ""
    A_Clipboard := text
    if !ClipWait(2)
        throw Error("클립보드 설정 실패")
    Keys("^a^v", tab)
}

main := StartOrConnect()
savedClip := ClipboardAll()                       ; 사용자의 클립보드 백업
WinActivate main
if !WinWaitActive(main, , 3)
    throw Error("창을 활성화하지 못했습니다.")
Keys("^n", false)                                 ; Ctrl+N : 새로 입력 (고객코드에 포커스)
Sleep 300

TypeText(rec["고객코드"])
PasteText(rec["고객명"])
TypeText(rec["전화번호"])
downs := IndexOf(REGIONS, rec["지역"]) - 1        ; 콤보: 첫 항목에서 ↓ 몇 번
Keys("{Home}{Down " Max(downs, 0) "}")
Keys(rec["고객구분"] = "법인" ? "{Right}" : "")   ; 포커스는 '개인' 라디오 → → 키로 '법인'
Keys(rec["VIP"] = "Y" ? "{Space}" : "")           ; 체크박스 토글
TypeText(rec["주문일자"])
TypeText(rec["수량"])
TypeText(rec["단가"])
PasteText(rec["비고"], false)

Sleep 1000
Send "!s"                                         ; Alt+S = 저장(&S)
box := WaitMsgBox(main, 20)
A_Clipboard := savedClip                          ; 클립보드 복원
if box {
    CloseMsgBox(box)
    MsgBox "[" box.title "] " box.text, "04 - 키보드 입력 결과"
} else
    MsgBox "메시지박스가 나타나지 않았습니다.", "04", "Icon!"
