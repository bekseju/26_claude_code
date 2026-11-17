; =============================================================================
; [2단계] 한 건 입력 - 컨트롤 직접 제어 (창이 뒤에 있어도 동작, 마우스/키보드 안 씀)
;
;   실제 입력은 common.ahk 의 FillOrder / SaveOrder 가 합니다. 안에서 쓰는 함수:
;   ControlSetText      : Edit 에 값 넣기 (WM_SETTEXT)
;   ControlChooseString : 콤보박스 항목 선택
;   ControlSetChecked   : 라디오/체크박스
;   ControlClick        : 버튼 클릭 (버튼 텍스트 또는 ClassNN 으로 지정)
;   WinWait #32770      : 메시지박스 대기
;   이 파일에서는 ControlGet* 로 입력 결과를 다시 읽어 확인합니다.
;
;   ※ 컨트롤 이름(ClassNN)은 'Window Spy' (AutoHotkey 트레이 메뉴) 로 확인할 수 있습니다.
;     VB6 예: ThunderRT6TextBox3,  Delphi 예: TEdit2
; =============================================================================
#Requires AutoHotkey v2.0
#Include common.ahk

rec := Map("고객코드", "C9101", "고객명", "AHK홍길동", "전화번호", "010-2222-0000", "지역", "광주"
         , "고객구분", "법인", "VIP", "Y", "주문일자", "2026-11-07", "수량", "6", "단가", "7000"
         , "비고", "AutoHotkey 한 건 입력 테스트")

main := StartOrConnect()
f := FillOrder(main, rec)

; 입력 결과 다시 읽기 (금액은 프로그램이 계산한 값)
Sleep 200
msg := "고객코드: " ControlGetText(f["고객코드"])
    . "`n지역: " ControlGetChoice(f["지역"])
    . "`n법인: " ControlGetChecked("법인", main) "  VIP: " ControlGetChecked("VIP", main)
    . "`n프로그램이 계산한 금액: " ControlGetText(f["금액"])
ToolTip msg
Sleep 1500
ToolTip

try
    result := SaveOrder(main)
catch as e
    result := "오류: " e.Message
MsgBox msg "`n`n결과: " result, "02 - 한 건 입력"
