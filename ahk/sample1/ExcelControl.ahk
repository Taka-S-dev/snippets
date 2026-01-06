; ==============================================================================
; Module:       ExcelControl.ahk
; Description:  Excel操作に関連する機能 (パレット操作ベース版)
; ==============================================================================

#Requires AutoHotkey v2.0

class ExcelControl {
    /**
     * 選択中の文字列を赤にする
     */
    static SetFontColorRed() {
        if !WinActive("ahk_class XLMAIN")
            return
        Send("{Alt}hfc")
        Sleep(150)
        Send("{Down 7}{Right 1}{Enter}")
    }

    /**
     * 選択中の文字列を黒(自動)に戻す
     */
    static SetFontColorBlack() {
        if !WinActive("ahk_class XLMAIN")
            return
        Send("{Alt}hfc")
        Sleep(150)
        Send("{Enter}")
    }

    /**
     * 取り消し線の切り替え (Strikethrough)
     */
    static SetFontColorStrikethrough() {
        if !WinActive("ahk_class XLMAIN")
            return
        ; ホーム(H) -> フォント設定(FN) を開き、取り消し線(Alt+K)をチェックして確定
        Send("{Alt}hfn")
        Sleep(200)
        Send("!k{Enter}")
    }
}
