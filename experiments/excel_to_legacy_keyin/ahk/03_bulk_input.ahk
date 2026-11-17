; =============================================================================
; [3단계] CSV 전체 일괄 입력 - 실무형 구조
;   - 사전 검증(CheckOrder) → 입력(FillOrder) → 저장·판정(SaveOrder)
;   - 행마다 결과를 data\orders_result_ahk.csv 에 즉시 기록 (중간에 멈춰도 보존)
;   - 재실행 시 '성공' 행은 건너뜀 (이어서 하기)
;   - Esc : 현재 건을 마치고 중지,  F8 : 일시정지/재개
;   - 로그: ahk\bulk_input_ahk.log
; =============================================================================
#Requires AutoHotkey v2.0
#SingleInstance Force
#Include common.ahk

stopRequested := false
~Esc:: {                       ; ~ : Esc 키를 가로채지 않고 다른 프로그램에도 전달
    global stopRequested := true
    ToolTip "중지 요청됨 - 현재 건을 마치고 멈춥니다."
}
F8::Pause -1

RESULT_CSV := DATA_DIR "\orders_result_ahk.csv"
LOG_FILE := A_ScriptDir "\bulk_input_ahk.log"

WriteLog(text) {
    FileAppend Timestamp() " " text "`n", LOG_FILE, "UTF-8"
    ToolTip text, 10, 10
}

WriteResult(rowNo, rec, result) {
    FileAppend Format('{1},{2},"{3}","{4}",{5}`n', rowNo, rec["고객코드"], rec["고객명"]
        , StrReplace(result, '"', '""'), Timestamp()), RESULT_CSV, "UTF-8"
}

; 사전 검증 → 입력 → 저장. 예외가 나면 화면을 정리하고 "오류: ..." 반환
ProcessOrder(main, rec) {
    if (reason := CheckOrder(rec))
        return "건너뜀: " reason
    try {
        FillOrder(main, rec)
        return SaveOrder(main)
    } catch as e {
        CloseAllMsgBoxes(main)
        return "오류: " e.Message
    }
}

; 이전 결과에서 '성공'한 행 번호 (이어서 하기)
doneRows := Map()
if FileExist(RESULT_CSV) {
    for r in ReadCsv(RESULT_CSV)
        if (r["처리결과"] = "성공")
            doneRows[r["행"]] := true
} else
    FileAppend "행,고객코드,고객명,처리결과,처리시각`n", RESULT_CSV, "UTF-8"

orders := ReadCsv(DATA_DIR "\orders.csv")
if MsgBox(orders.Length "건 중 " (orders.Length - doneRows.Count) "건을 입력합니다.`n"
        . "실행 중 Esc = 중지, F8 = 일시정지`n계속할까요?", "03 - 일괄 입력", "OKCancel") != "OK"
    ExitApp

main := StartOrConnect()
okCount := 0, failCount := 0
for i, rec in orders {
    rowNo := String(i + 1)                                  ; CSV 행 번호 (헤더 = 1행)
    if doneRows.Has(rowNo)
        continue
    if stopRequested {
        WriteLog("Esc 로 중지")
        break
    }
    result := ProcessOrder(main, rec)
    (result = "성공") ? okCount++ : failCount++
    WriteLog(Format("({1}/{2}) {3} {4} → {5}", i, orders.Length, rec["고객코드"], rec["고객명"], result))
    WriteResult(rowNo, rec, result)
}
ToolTip
MsgBox "완료: 성공 " okCount " / 실패·건너뜀 " failCount "`n결과: " RESULT_CSV, "03 - 일괄 입력"
ExitApp
