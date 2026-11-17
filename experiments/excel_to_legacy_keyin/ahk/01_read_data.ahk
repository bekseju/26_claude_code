; =============================================================================
; [1단계] 데이터 읽기 - 두 가지 방법
;   A) CSV 읽기        : Excel 설치 불필요, 빠름, 가장 안정적  (엑셀에서 "다른 이름으로 저장 → CSV UTF-8")
;   B) Excel COM 읽기  : xlsx 를 직접 읽음 (Excel 설치 필요)
;        .Text  = 화면에 '보이는' 문자열 (날짜 서식 그대로 "2026-11-01")  ← 키인 용도로 추천
;        .Value = 서식이 적용되지 않은 원본 값 (숫자 서식·앞자리 0 등이 화면과 다를 수 있음)
; =============================================================================
#Requires AutoHotkey v2.0
#Include common.ahk

; ---- A) CSV --------------------------------------------------------------
rows := ReadCsv(DATA_DIR "\orders.csv")
out := "[CSV] " rows.Length "건`n"
for r in rows
    out .= Format("{1} | {2} | {3} | {4} | 수량={5}`n", r["고객코드"], r["고객명"], r["지역"], r["주문일자"], r["수량"])
MsgBox out, "01 - CSV 읽기"

; ---- B) Excel COM --------------------------------------------------------
xl := ComObject("Excel.Application")           ; 이미 열린 엑셀을 쓰려면: xl := ComObjActive("Excel.Application")
xl.Visible := false
xl.DisplayAlerts := false
try {
    wb := xl.Workbooks.Open(DATA_DIR "\orders.xlsx", 0, true)   ; (경로, 링크갱신 안함, 읽기전용)
    ws := wb.Worksheets("주문데이터")
    lastRow := ws.Cells(ws.Rows.Count, 1).End(-4162).Row        ; -4162 = xlUp : A열 마지막 데이터 행

    ; 헤더 이름 → 열 번호 (열 순서가 바뀌어도 동작)
    col := Map()
    loop ws.UsedRange.Columns.Count
        col[ws.Cells(1, A_Index).Text] := A_Index

    out := "[Excel COM] " (lastRow - 1) "건`n"
    loop lastRow - 1 {
        r := A_Index + 1
        out .= Format("행{1} | {2} | {3} | 전화(.Text)={4} | 날짜(.Text)={5} | 날짜(.Value)={6}`n"
            , r
            , Trim(ws.Cells(r, col["고객코드"]).Text)
            , Trim(ws.Cells(r, col["고객명"]).Text)
            , ws.Cells(r, col["전화번호"]).Text
            , ws.Cells(r, col["주문일자"]).Text
            , ws.Cells(r, col["주문일자"]).Value)
    }
    wb.Close(false)
} finally {
    xl.Quit()                                   ; 꼭 종료! 안 하면 EXCEL.EXE 가 백그라운드에 남음
    xl := ""
}
MsgBox out, "01 - Excel COM 읽기"
