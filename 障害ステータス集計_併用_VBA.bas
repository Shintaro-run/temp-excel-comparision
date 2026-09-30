Attribute VB_Name = "障害ステータス集計"
Option Explicit

' ============================================================
'  障害ステータス集計（親 × 子）／併用版
'  ----------------------------------------------------------
'  実行すると2つを出します。
'   (1) 「ステータス」シートに、クロス集計表＋Excelの積み上げ縦棒（静的）。
'   (2) このブックと同じフォルダに、触れる・動く HTML ダッシュボードを書き出し、開く。
'       （棒が1本ずつ伸びる／ホバーで件数／棒・表クリックで親を掘り下げ／凡例クリックで子を絞り込み／件数・%切替）
'
'  ★ステータス名は J列・L列から自動で拾います（固定の一覧は持ちません）。
'    名前が変わっても、初期設定なしで表・グラフに反映されます。
'  ★集計する行は F列に文字列がある行だけです。F列が空の行は完全に無視します。
'
'  使い方：Alt+F11 → 標準モジュールに貼り付け → マクロ「ステータス集計を作る」を実行。
'  ⚠️ まず下の【設定】のシート名・列・開始行だけ、実物に合っているか確認してください。
'  ⚠️ HTML の書き出しには、参照設定は不要です（ADODB を遅延バインドで使います）。
' ============================================================

' ===== 設定（環境に合わせてここだけ）=====
Private Const SHEET_DATA As String = "障害一覧"
Private Const FIRST_ROW  As Long = 3
Private Const COL_ID As String = "A"      ' 障害ID（参考・集計の判定には使わない）
Private Const COL_KEEP As String = "F"     ' この列に文字列がある行だけ集計する（空なら完全に無視）
Private Const COL_AT As String = "J"      ' 親：親ステータス
Private Const COL_NS As String = "L"      ' 子：子ステータス
Private Const OUT_SHEET As String = "ステータス"
Private Const HTML_NAME As String = "障害ステータス_ダッシュボード.html"
' 表記ゆれをまとめたいとき True（既定 False＝J/Lの値をそのまま使う。名前が変わっても追随）
Private Const USE_NORMALIZE As Boolean = False

Sub ステータス集計を作る()
    Dim wsD As Worksheet, wsO As Worksheet
    On Error Resume Next
    Set wsD = ThisWorkbook.Worksheets(SHEET_DATA)
    On Error GoTo 0
    If wsD Is Nothing Then MsgBox "シート「" & SHEET_DATA & "」が見つかりません。", vbExclamation: Exit Sub

    Dim cKeep As Long, cAT As Long, cNS As Long, cID As Long
    cKeep = wsD.Range(COL_KEEP & "1").Column
    cAT = wsD.Range(COL_AT & "1").Column
    cNS = wsD.Range(COL_NS & "1").Column
    cID = wsD.Range(COL_ID & "1").Column
    Dim lastRow As Long
    lastRow = wsD.Cells(wsD.Rows.Count, cKeep).End(xlUp).Row   ' F列の最終行までを見る
    If lastRow < FIRST_ROW Then MsgBox "データ行がありません（" & COL_KEEP & "列が空です）。", vbExclamation: Exit Sub

    ' ---- クロス集計（親 × 子）。ステータスは自動で拾う ----
    Dim tbl As Object, atKeys As Object, nsKeys As Object, ids As Object
    Set tbl = CreateObject("Scripting.Dictionary")
    Set atKeys = CreateObject("Scripting.Dictionary")
    Set nsKeys = CreateObject("Scripting.Dictionary")
    Set ids = CreateObject("Scripting.Dictionary")   ' a -> (n -> IDをChr(1)でつないだ文字列)
    Dim r As Long, a As String, n As String, total As Long
    For r = FIRST_ROW To lastRow
        If Trim(CStr(wsD.Cells(r, cKeep).Value)) <> "" Then   ' F列に文字列がある行だけ集計
            a = 値(wsD.Cells(r, cAT).Value, True)
            n = 値(wsD.Cells(r, cNS).Value, False)
            If Not atKeys.Exists(a) Then atKeys.Add a, atKeys.Count
            If Not nsKeys.Exists(n) Then nsKeys.Add n, nsKeys.Count
            If Not tbl.Exists(a) Then tbl.Add a, CreateObject("Scripting.Dictionary")
            Dim rw As Object: Set rw = tbl(a)
            If Not rw.Exists(n) Then rw.Add n, 0
            rw(n) = rw(n) + 1
            total = total + 1
            ' セルごとの障害ID（A列。空なら行番号）を集める
            Dim theId As String: theId = Trim(CStr(wsD.Cells(r, cID).Value))
            If theId = "" Then theId = "行" & r
            If Not ids.Exists(a) Then ids.Add a, CreateObject("Scripting.Dictionary")
            Dim rid As Object: Set rid = ids(a)
            If Not rid.Exists(n) Then rid.Add n, ""
            rid(n) = rid(n) & IIf(rid(n) = "", "", Chr(1)) & theId
        End If
    Next r
    If total = 0 Then MsgBox "対象の障害が0件でした。", vbExclamation: Exit Sub

    Dim atArr As Variant, nsArr As Variant
    atArr = atKeys.Keys
    nsArr = 並べ替えNS(nsKeys.Keys)

    ' ---- (1) 「ステータス」シート：表＋Excel積み上げグラフ ----
    Application.DisplayAlerts = False
    On Error Resume Next: ThisWorkbook.Worksheets(OUT_SHEET).Delete: On Error GoTo 0
    Application.DisplayAlerts = True
    Set wsO = ThisWorkbook.Worksheets.Add(After:=wsD): wsO.Name = OUT_SHEET

    Dim i As Long, j As Long, rr As Long, colTot As Long, rowSum As Long
    Dim colSum() As Long: ReDim colSum(0 To UBound(nsArr))
    colTot = 2 + (UBound(nsArr) + 1)
    wsO.Cells(1, 1).Value = "親ステータス ＼ 子ステータス"
    wsO.Cells(1, 1).Font.Bold = True
    For j = 0 To UBound(nsArr)
        wsO.Cells(1, 2 + j).Value = nsArr(j): wsO.Cells(1, 2 + j).Font.Bold = True
        wsO.Cells(1, 2 + j).HorizontalAlignment = xlCenter
    Next j
    wsO.Cells(1, colTot).Value = "合計": wsO.Cells(1, colTot).Font.Bold = True
    For i = 0 To UBound(atArr)
        rr = 2 + i: wsO.Cells(rr, 1).Value = atArr(i): wsO.Cells(rr, 1).Font.Bold = True
        rowSum = 0
        Dim rw2 As Object: Set rw2 = tbl(atArr(i))
        For j = 0 To UBound(nsArr)
            Dim v As Long: v = 0
            If rw2.Exists(nsArr(j)) Then v = rw2(nsArr(j))
            wsO.Cells(rr, 2 + j).Value = v: rowSum = rowSum + v: colSum(j) = colSum(j) + v
        Next j
        wsO.Cells(rr, colTot).Value = rowSum: wsO.Cells(rr, colTot).Font.Bold = True
    Next i
    Dim totRow As Long: totRow = 2 + (UBound(atArr) + 1)
    wsO.Cells(totRow, 1).Value = "合計": wsO.Cells(totRow, 1).Font.Bold = True
    For j = 0 To UBound(nsArr)
        wsO.Cells(totRow, 2 + j).Value = colSum(j): wsO.Cells(totRow, 2 + j).Font.Bold = True
    Next j
    wsO.Cells(totRow, colTot).Value = total: wsO.Cells(totRow, colTot).Font.Bold = True
    wsO.Range(wsO.Cells(1, 1), wsO.Cells(totRow, colTot)).Borders.LineStyle = xlContinuous
    wsO.Columns("A").ColumnWidth = 30

    Dim ch As ChartObject
    Set ch = wsO.ChartObjects.Add(Left:=wsO.Cells(1, colTot + 2).Left, Top:=8, Width:=560, Height:=320)
    With ch.Chart
        .ChartType = xlColumnStacked
        .SetSourceData Source:=wsO.Range(wsO.Cells(1, 1), wsO.Cells(1 + (UBound(atArr) + 1), 1 + (UBound(nsArr) + 1)))
        .HasTitle = True: .ChartTitle.Text = "親ステータスごとの、子ステータスの内訳"
        On Error Resume Next
        .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.Text = "件数"
        .HasLegend = True: .Legend.Position = xlLegendPositionBottom
        On Error GoTo 0
    End With

    ' ---- (2) HTML ダッシュボードを書き出して開く ----
    Dim jsonS As String: jsonS = JSON(atArr, nsArr, tbl, colSum, total, ids)
    Dim htmlS As String: htmlS = Replace(HTMLテンプレ(), "{{DATA}}", jsonS)
    Dim path As String: path = HTML保存先()
    書き出しUTF8 htmlS, path
    wsO.Activate: wsO.Range("A1").Select
    On Error Resume Next
    ThisWorkbook.FollowHyperlink path
    On Error GoTo 0
    MsgBox "集計しました（対象 " & total & " 件）。" & vbCrLf & _
           "・「" & OUT_SHEET & "」シートに表とグラフ" & vbCrLf & _
           "・ダッシュボード：" & path, vbInformation
End Sub

' 値の取り出し（USE_NORMALIZE=False ならそのまま。空欄は（未設定）に）
Private Function 値(ByVal x As Variant, ByVal isAT As Boolean) As String
    Dim t As String: t = Trim(CStr(x))
    If t = "" Then 値 = "（未設定）": Exit Function
    If Not USE_NORMALIZE Or isAT Then 値 = t: Exit Function
    値 = 正規化NS(t)
End Function

' 表記ゆれのまとめ（USE_NORMALIZE=True のときだけ使う。実際の値に合わせて増やす）
Private Function 正規化NS(ByVal t As String) As String
    Dim low As String: low = LCase(t)
    Select Case True
        Case InStr(t, "終了") > 0 Or InStr(t, "完了") > 0 Or InStr(t, "クローズ") > 0 Or low = "closed" Or low = "done": 正規化NS = "終了"
        Case InStr(t, "計画") > 0: 正規化NS = "対応計画済"
        Case InStr(t, "対応中") > 0 Or InStr(t, "作業中") > 0 Or InStr(t, "進行") > 0 Or low = "in progress": 正規化NS = "対応中"
        Case InStr(t, "対象外") > 0 Or InStr(t, "範囲外") > 0 Or InStr(t, "無関係") > 0: 正規化NS = "対象外(子)"
        Case InStr(t, "保留") > 0 Or InStr(t, "確認待") > 0 Or InStr(t, "待ち") > 0: 正規化NS = "保留・確認待ち"
        Case Else: 正規化NS = t
    End Select
End Function

' 見やすい順に子を並べる（決まったものを前に、未知はそのあと）
Private Function 並べ替えNS(ByVal keys As Variant) As Variant
    Dim pref As Variant: pref = Array("終了", "対応計画済", "対応中", "保留・確認待ち", "対象外(子)", "（未設定）")
    Dim out() As String, c As Long: ReDim out(0 To UBound(keys))
    Dim p As Variant, k As Variant
    For Each p In pref
        For Each k In keys
            If k = p Then out(c) = k: c = c + 1
        Next k
    Next p
    For Each k In keys
        Dim found As Boolean: found = False
        For Each p In pref
            If k = p Then found = True: Exit For
        Next p
        If Not found Then out(c) = k: c = c + 1
    Next k
    並べ替えNS = out
End Function

' クロス集計を JSON 文字列に
Private Function JSON(atArr As Variant, nsArr As Variant, tbl As Object, colSum() As Long, total As Long, ids As Object) As String
    Dim s As String, i As Long, j As Long
    s = "{""at"":["
    For i = 0 To UBound(atArr): s = s & IIf(i > 0, ",", "") & JQ(CStr(atArr(i))): Next i
    s = s & "],""ns"":["
    For j = 0 To UBound(nsArr): s = s & IIf(j > 0, ",", "") & JQ(CStr(nsArr(j))): Next j
    s = s & "],""tbl"":{"
    For i = 0 To UBound(atArr)
        s = s & IIf(i > 0, ",", "") & JQ(CStr(atArr(i))) & ":{"
        Dim rw As Object: Set rw = tbl(atArr(i))
        For j = 0 To UBound(nsArr)
            Dim v As Long: v = 0
            If rw.Exists(nsArr(j)) Then v = rw(nsArr(j))
            s = s & IIf(j > 0, ",", "") & JQ(CStr(nsArr(j))) & ":" & v
        Next j
        s = s & "}"
    Next i
    s = s & "},""colsum"":["
    For j = 0 To UBound(nsArr): s = s & IIf(j > 0, ",", "") & colSum(j): Next j
    s = s & "],""total"":" & total
    ' セルごとの障害ID（空でないセルだけ出す）
    s = s & ",""ids"":{"
    Dim firstA As Boolean: firstA = True
    Dim rid2 As Object, firstN As Boolean, parts() As String, k As Long
    For i = 0 To UBound(atArr)
        If ids.Exists(atArr(i)) Then
            Set rid2 = ids(atArr(i))
            If Not firstA Then s = s & ","
            firstA = False
            s = s & JQ(CStr(atArr(i))) & ":{"
            firstN = True
            For j = 0 To UBound(nsArr)
                If rid2.Exists(nsArr(j)) Then
                    If Not firstN Then s = s & ","
                    firstN = False
                    parts = Split(rid2(nsArr(j)), Chr(1))
                    s = s & JQ(CStr(nsArr(j))) & ":["
                    For k = 0 To UBound(parts): s = s & IIf(k > 0, ",", "") & JQ(parts(k)): Next k
                    s = s & "]"
                End If
            Next j
            s = s & "}"
        End If
    Next i
    s = s & "}}"
    JSON = s
End Function

Private Function JQ(ByVal s As String) As String
    s = Replace(s, "\", "\\"): s = Replace(s, """", "\""")
    s = Replace(s, vbCr, " "): s = Replace(s, vbLf, " ")
    JQ = """" & s & """"
End Function

Private Function HTML保存先() As String
    Dim p As String: p = ThisWorkbook.path
    If p = "" Then p = Environ$("TEMP")   ' 未保存のブックなら一時フォルダ
    If Right$(p, 1) <> Application.PathSeparator Then p = p & Application.PathSeparator
    HTML保存先 = p & HTML_NAME
End Function

' UTF-8 で書き出す（日本語が化けないように。ADODB は遅延バインド）
Private Sub 書き出しUTF8(ByVal text As String, ByVal path As String)
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "UTF-8": st.Open
    st.WriteText text
    st.SaveToFile path, 2   ' 2 = 上書き
    st.Close
End Sub

Private Function HTMLテンプレ() As String
    Dim s As String
    s = s & "<!DOCTYPE html><html lang=ja><head><meta charset=utf-8>" & vbCrLf
    s = s & "<meta name=viewport content=""width=device-width, initial-scale=1"">" & vbCrLf
    s = s & "<title>障害ステータス ダッシュボード</title>" & vbCrLf
    s = s & "<style>" & vbCrLf
    s = s & ":root{--bg:#f2f5f9;--card:#fff;--ink:#1c2430;--sub:#5a6675;--line:#dbe2ea;--shadow:0 2px 12px rgba(20,40,80,.08)}" & vbCrLf
    s = s & ":root[data-theme=dark]{--bg:#0f141b;--card:#1a212b;--ink:#e6ecf3;--sub:#9aa7b6;--line:#33404f;--shadow:0 2px 12px rgba(0,0,0,.45)}" & vbCrLf
    s = s & ":root[data-theme=dark] thead th{background:#243244}" & vbCrLf
    s = s & ":root[data-theme=dark] .rowh{background:#1f2a38}" & vbCrLf
    s = s & ":root[data-theme=dark] .tot,:root[data-theme=dark] tfoot td,:root[data-theme=dark] tfoot th{background:#202a34}" & vbCrLf
    s = s & ":root[data-theme=dark] .seg-ctl{background:#243244}" & vbCrLf
    s = s & ":root[data-theme=dark] .boxes{background:#131a22}" & vbCrLf
    s = s & ":root[data-theme=dark] .applybtn{background:#1f2a38}" & vbCrLf
    s = s & ":root[data-theme=dark] .detail .chip,:root[data-theme=dark] .mtag,:root[data-theme=dark] .idchip{background:#243244}" & vbCrLf
    s = s & ":root[data-theme=dark] .lg:hover,:root[data-theme=dark] tbody tr:hover{background:#243244}" & vbCrLf
    s = s & ":root[data-theme=dark] .lg[aria-pressed=true],:root[data-theme=dark] tbody tr.on,:root[data-theme=dark] th.oncol,:root[data-theme=dark] td.oncol{background:#2a3a52}" & vbCrLf
    s = s & ":root[data-theme=dark] .seg-ctl button[aria-pressed=true],:root[data-theme=dark] #viewToggle[aria-pressed=true],:root[data-theme=dark] #darkBtn[aria-pressed=true],:root[data-theme=dark] #cbBtn[aria-pressed=true],:root[data-theme=dark] #mergeToggle[aria-pressed=true]{background:#2a3a52;color:#cfe0ff;border-color:#3f6fd0}" & vbCrLf
    s = s & "*{box-sizing:border-box}" & vbCrLf
    s = s & "body{margin:0;background:var(--bg);color:var(--ink);font-family:""Hiragino Kaku Gothic ProN"",""Helvetica Neue"",Arial,sans-serif}" & vbCrLf
    s = s & ".wrap{max-width:1000px;margin:0 auto;padding:22px}" & vbCrLf
    s = s & "h1{font-size:19px;margin:0 0 3px}.sub{font-size:12.5px;color:var(--sub);margin:0 0 16px}" & vbCrLf
    s = s & ".bar-top{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:12px}" & vbCrLf
    s = s & ".seg-ctl{display:flex;background:#e7ecf3;border-radius:999px;padding:3px}" & vbCrLf
    s = s & ".seg-ctl button{border:0;background:transparent;font:inherit;font-size:13px;padding:6px 16px;border-radius:999px;cursor:pointer;color:var(--sub)}" & vbCrLf
    s = s & ".seg-ctl button[aria-pressed=true]{background:#2f5bd0;color:#fff;font-weight:700}" & vbCrLf
    s = s & ".legend{display:flex;gap:8px;flex-wrap:wrap;font-size:12.5px}" & vbCrLf
    s = s & ".lg{display:flex;align-items:center;gap:6px;cursor:pointer;padding:4px 9px;border-radius:999px;border:1px solid transparent;user-select:none;max-width:100%;word-break:break-word}" & vbCrLf
    s = s & ".lg:hover{background:#eef2f8}.lg i{width:13px;height:13px;border-radius:4px;flex-shrink:0}" & vbCrLf
    s = s & ".lg[aria-pressed=true]{background:#eaf1ff;border-color:#2f5bd0;font-weight:700}" & vbCrLf
    s = s & ".clearbtn{border:1px solid var(--line);background:#fff;font:inherit;font-size:12px;padding:6px 13px;border-radius:999px;cursor:pointer;color:var(--sub)}" & vbCrLf
    s = s & "#mergeToggle[aria-pressed=true]{background:#eaf1ff;border-color:#2f5bd0;color:#2f5bd0;font-weight:700}" & vbCrLf
    s = s & ".mergebox{background:var(--card);border:1px solid var(--line);border-radius:14px;box-shadow:var(--shadow);padding:14px 16px;margin-bottom:14px}" & vbCrLf
    s = s & ".mergecols{display:flex;gap:20px;flex-wrap:wrap}" & vbCrLf
    s = s & ".mergecol{flex:1;min-width:240px}" & vbCrLf
    s = s & ".mh{font-weight:700;font-size:13px;margin-bottom:8px}" & vbCrLf
    s = s & ".boxes{display:flex;flex-direction:column;gap:3px;max-height:170px;overflow:auto;border:1px solid var(--line);border-radius:8px;padding:8px;background:#fbfcfe}" & vbCrLf
    s = s & ".ck{display:flex;align-items:center;gap:7px;font-size:12.5px;cursor:pointer;padding:2px 3px;word-break:break-word}" & vbCrLf
    s = s & ".ck input{flex-shrink:0}.ck i{width:11px;height:11px;border-radius:3px;flex-shrink:0}" & vbCrLf
    s = s & ".mrow{display:flex;gap:8px;margin-top:8px}" & vbCrLf
    s = s & ".mrow input{flex:1;min-width:0;font:inherit;font-size:12.5px;padding:6px 9px;border:1px solid var(--line);border-radius:8px}" & vbCrLf
    s = s & ".mbtn{border:0;background:#2f5bd0;color:#fff;font:inherit;font-size:12.5px;font-weight:700;padding:6px 14px;border-radius:8px;cursor:pointer;flex-shrink:0}" & vbCrLf
    s = s & ".msg{color:#b0402f;font-size:12px;min-height:16px;margin-top:4px}" & vbCrLf
    s = s & ".msg.ok{color:#1f7a3d}" & vbCrLf
    s = s & ".mtags{display:flex;flex-wrap:wrap;gap:6px;margin-top:6px}" & vbCrLf
    s = s & ".mtag{display:inline-flex;align-items:center;gap:6px;background:#eef2f8;border-radius:999px;padding:3px 6px 3px 11px;font-size:12px;max-width:100%;word-break:break-word}" & vbCrLf
    s = s & ".mx{border:0;background:#c9d3e0;color:#333;border-radius:50%;width:18px;height:18px;line-height:1;cursor:pointer;font-size:12px;flex-shrink:0}" & vbCrLf
    s = s & ".mfoot{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-top:12px}" & vbCrLf
    s = s & ".hint{font-size:11.5px;color:var(--sub)}" & vbCrLf
    s = s & ".muted{color:var(--sub);font-size:12px}" & vbCrLf
    s = s & ".ic{width:14px;height:14px;vertical-align:-2px;margin-right:5px;fill:none;stroke:currentColor;stroke-width:2;stroke-linejoin:round}" & vbCrLf
    s = s & ".plist{margin-top:8px;display:flex;flex-direction:column;gap:6px}" & vbCrLf
    s = s & ".prow{display:flex;align-items:center;gap:8px}" & vbCrLf
    s = s & ".applybtn{flex:1;text-align:left;border:1px solid var(--line);background:#f7faff;font:inherit;font-size:12.5px;padding:7px 12px;border-radius:8px;cursor:pointer;color:var(--ink);word-break:break-word}" & vbCrLf
    s = s & ".applybtn:hover{background:#eaf1ff;border-color:#2f5bd0}" & vbCrLf
    s = s & ".pcount{font-size:11.5px;color:var(--sub);flex-shrink:0}" & vbCrLf
    s = s & "#impLabel{cursor:pointer;margin:0}" & vbCrLf
    s = s & ".vrow{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:9px}" & vbCrLf
    s = s & ".vlab{font-size:12.5px;font-weight:700;min-width:150px}" & vbCrLf
    s = s & "#viewPanel input,#viewPanel select{font:inherit;font-size:12.5px;padding:5px 9px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--ink)}" & vbCrLf
    s = s & "#viewToggle[aria-pressed=true],#darkBtn[aria-pressed=true],#cbBtn[aria-pressed=true]{background:#eaf1ff;border-color:#2f5bd0;color:#2f5bd0;font-weight:700}" & vbCrLf
    s = s & ".cellhead{display:flex;align-items:center;justify-content:space-between;gap:10px;font-weight:700;font-size:13px;margin-bottom:8px}" & vbCrLf
    s = s & ".cellids{display:flex;flex-wrap:wrap;gap:6px}" & vbCrLf
    s = s & ".idchip{background:#eef2f8;border-radius:6px;padding:3px 9px;font-size:12px;font-family:""SFMono-Regular"",Menlo,Consolas,monospace}" & vbCrLf
    s = s & ".idcell{cursor:pointer}tbody tr:hover .idcell,.idcell:hover{outline:1px solid #b9c7de;outline-offset:-1px}" & vbCrLf
    s = s & "#toast{position:fixed;left:50%;bottom:24px;transform:translateX(-50%);background:#1c2430;color:#fff;font-size:12.5px;padding:8px 16px;border-radius:8px;opacity:0;transition:opacity .2s;z-index:20;pointer-events:none;box-shadow:0 4px 14px rgba(0,0,0,.3)}" & vbCrLf
    s = s & ".printhead{display:none}" & vbCrLf
    s = s & ".card{background:var(--card);border:1px solid var(--line);border-radius:16px;box-shadow:var(--shadow);padding:20px 22px}" & vbCrLf
    s = s & ".chart{display:flex;gap:26px;align-items:flex-end;height:300px;padding:8px 4px 0}" & vbCrLf
    s = s & ".col{display:flex;flex-direction:column;align-items:center;cursor:pointer;flex:1;min-width:0}" & vbCrLf
    s = s & ".stack{display:flex;flex-direction:column-reverse;width:74px;max-width:80%;border-bottom:2px solid #c2ccd8;position:relative}" & vbCrLf
    s = s & ".seg{width:100%;color:#fff;font-size:11px;font-weight:700;display:flex;align-items:center;justify-content:center;overflow:hidden;" & vbCrLf
    s = s & "     height:0;transition:height .5s cubic-bezier(.22,.61,.36,1),opacity .2s;opacity:.96}" & vbCrLf
    s = s & ".seg:hover{filter:brightness(1.08)}.seg.mute{opacity:.14}" & vbCrLf
    s = s & ".xl{margin-top:9px;text-align:center;font-size:12.5px;color:#33445a;width:100%;padding:0 2px}" & vbCrLf
    s = s & ".xl .nm{display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;word-break:break-word;line-height:1.25;min-height:2.5em}" & vbCrLf
    s = s & ".xl b{display:block;font-size:15px;color:var(--ink);margin-top:2px}" & vbCrLf
    s = s & ".col.dim{opacity:.3;transition:opacity .2s}" & vbCrLf
    s = s & ".col.on .stack{outline:2px solid #2f5bd0;outline-offset:3px;border-radius:5px}" & vbCrLf
    s = s & "#splitBtn[aria-pressed=true]{background:#eaf1ff;border-color:#2f5bd0;color:#2f5bd0;font-weight:700}" & vbCrLf
    s = s & ":root[data-theme=dark] #splitBtn[aria-pressed=true]{background:#2a3a52;color:#cfe0ff;border-color:#3f6fd0}" & vbCrLf
    s = s & ".chart.split{gap:0}" & vbCrLf
    s = s & ".splitgroup{flex:1;min-width:0;display:flex;flex-direction:column}" & vbCrLf
    s = s & ".splitgroup .ghd{text-align:center;font-size:12px;font-weight:700;color:var(--sub);margin-bottom:6px;padding-bottom:4px;border-bottom:2px solid var(--line)}" & vbCrLf
    s = s & ".splitgroup .gbars{flex:1;display:flex;gap:18px;align-items:flex-end}" & vbCrLf
    s = s & ".split-div{width:1px;background:var(--line);margin:0 18px;align-self:stretch}" & vbCrLf
    s = s & ".stack.solid{border-radius:4px 4px 0 0}" & vbCrLf
    s = s & ".col[draggable=true]{cursor:grab}" & vbCrLf
    s = s & ".col.dragging{opacity:.4}" & vbCrLf
    s = s & ".col.dropok .stack{outline:2px dashed #2f5bd0;outline-offset:3px;border-radius:5px}" & vbCrLf
    s = s & "#tip{position:fixed;pointer-events:none;background:#1c2430;color:#fff;font-size:12px;padding:6px 9px;border-radius:7px;" & vbCrLf
    s = s & "     opacity:0;transition:opacity .12s;z-index:9;white-space:normal;max-width:240px;box-shadow:0 4px 14px rgba(0,0,0,.25)}" & vbCrLf
    s = s & "#tip b{color:#ffd98a}" & vbCrLf
    s = s & ".detail{margin-top:14px;font-size:13px;color:var(--sub);min-height:20px}" & vbCrLf
    s = s & ".detail .chip{display:inline-flex;align-items:center;gap:5px;background:#f0f3f8;border-radius:999px;padding:3px 10px;margin:3px 5px 0 0;color:var(--ink);max-width:100%;word-break:break-word}" & vbCrLf
    s = s & ".detail .chip i{width:10px;height:10px;border-radius:3px;flex-shrink:0}" & vbCrLf
    s = s & ".tablewrap{margin-top:18px;overflow-x:auto;-webkit-overflow-scrolling:touch;border:1px solid var(--line);border-radius:10px}" & vbCrLf
    s = s & "table{border-collapse:collapse;font-size:13px;width:auto;min-width:100%}" & vbCrLf
    s = s & "th,td{border:1px solid var(--line);padding:7px 10px;text-align:center;white-space:normal;word-break:break-word;vertical-align:middle}" & vbCrLf
    s = s & "thead th{background:#eaf0fb;font-weight:700;max-width:150px;min-width:58px}" & vbCrLf
    s = s & ".rowh{background:#f4f7fb;text-align:left;font-weight:700;position:sticky;left:0;z-index:2;max-width:180px;min-width:96px}" & vbCrLf
    s = s & "thead .rowh{z-index:3}" & vbCrLf
    s = s & "tbody tr{cursor:pointer}tbody tr:hover{background:#f7faff}tbody tr.on{background:#eaf1ff}" & vbCrLf
    s = s & "th.oncol,td.oncol{background:#eaf1ff}" & vbCrLf
    s = s & ".tot{background:#f6f8fb;font-weight:700}tfoot td,tfoot th{background:#f4f4f4;font-weight:700}" & vbCrLf
    s = s & "@media(prefers-reduced-motion:reduce){.seg{transition:opacity .2s}}" & vbCrLf
    s = s & "@media(max-width:620px){.chart{gap:14px;height:250px}.stack{width:auto}.mrow{flex-wrap:wrap}.mrow input{width:100%}.mbtn{width:100%}}" & vbCrLf
    s = s & "@media print{" & vbCrLf
    s = s & "  @page{size:A4 landscape;margin:12mm}" & vbCrLf
    s = s & "  body{background:#fff}" & vbCrLf
    s = s & "  .wrap{max-width:none;padding:0}" & vbCrLf
    s = s & "  .seg-ctl,#splitBtn,#mergeToggle,#saveBtn,#loadBtn,#pdfBtn,#csvBtn,#viewToggle,#clear,#mergePanel,#presetPanel,#viewPanel,#cellbox,#toast,#tip,.detail{display:none!important}" & vbCrLf
    s = s & "  .bar-top{margin:0 0 8px}" & vbCrLf
    s = s & "  .printhead{display:block;margin:0 0 10px}" & vbCrLf
    s = s & "  .printhead .ph-title{font-size:15pt;font-weight:700}" & vbCrLf
    s = s & "  .printhead .ph-sub{font-size:9.5pt;color:#444;margin-top:2px}" & vbCrLf
    s = s & "  h1{font-size:16pt;margin:0 0 4px}.sub{font-size:9pt;margin:0 0 10px}" & vbCrLf
    s = s & "  .legend{font-size:9pt;gap:6px}" & vbCrLf
    s = s & "  .card{box-shadow:none;border:1px solid #bbb;padding:12px 12px 8px;margin-bottom:12px;break-inside:avoid}" & vbCrLf
    s = s & "  .chart{height:300px;gap:18px;padding-top:4px}" & vbCrLf
    s = s & "  .tablewrap{overflow:visible!important;border:1px solid #bbb;border-radius:0;margin-top:0}" & vbCrLf
    s = s & "  table{width:100%;font-size:8.5pt}" & vbCrLf
    s = s & "  th,td{padding:3px 5px}" & vbCrLf
    s = s & "  thead th{max-width:none;min-width:0}.rowh{position:static;max-width:none;min-width:0}" & vbCrLf
    s = s & "  thead{display:table-header-group}" & vbCrLf
    s = s & "  tr{break-inside:avoid}" & vbCrLf
    s = s & "  *{-webkit-print-color-adjust:exact;print-color-adjust:exact}" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "</style></head><body>" & vbCrLf
    s = s & "<div id=tip></div>" & vbCrLf
    s = s & "<div id=toast></div>" & vbCrLf
    s = s & "<div class=wrap>" & vbCrLf
    s = s & "<div id=printHead class=printhead></div>" & vbCrLf
    s = s & "<h1>障害ステータス ダッシュボード</h1>" & vbCrLf
    s = s & "<p class=sub>親ステータスごとに、子ステータスを積み上げ。棒・表・凡例で掘り下げます。<span id=meta></span></p>" & vbCrLf
    s = s & "<div class=bar-top>" & vbCrLf
    s = s & "  <div class=seg-ctl role=group aria-label=表示>" & vbCrLf
    s = s & "    <button data-m=count aria-pressed=true>件数</button>" & vbCrLf
    s = s & "    <button data-m=pct aria-pressed=false>割合 %</button>" & vbCrLf
    s = s & "  </div>" & vbCrLf
    s = s & "  <button class=clearbtn id=splitBtn aria-pressed=false>親子で左右に分ける</button>" & vbCrLf
    s = s & "  <button class=clearbtn id=mergeToggle aria-pressed=false>ステータスをまとめる</button>" & vbCrLf
    s = s & "  <button class=clearbtn id=saveBtn title=""いまのまとめをルールとして保存""><svg class=ic viewBox=""0 0 24 24""><path d=""M4 4h13l3 3v13H4z""/><path d=""M8 4v5h7""/><rect x=""8"" y=""13"" width=""8"" height=""7""/></svg>保存</button>" & vbCrLf
    s = s & "  <button class=clearbtn id=loadBtn title=""保存したルールを呼び出す""><svg class=ic viewBox=""0 0 24 24""><path d=""M3 6h6l2 2h10v11H3z""/></svg>呼び出し</button>" & vbCrLf
    s = s & "  <button class=clearbtn id=pdfBtn>PDF出力（A4）</button>" & vbCrLf
    s = s & "  <button class=clearbtn id=csvBtn>CSVコピー</button>" & vbCrLf
    s = s & "  <button class=clearbtn id=viewToggle aria-pressed=false>表示設定</button>" & vbCrLf
    s = s & "  <button class=clearbtn id=clear hidden>絞り込み解除</button>" & vbCrLf
    s = s & "</div>" & vbCrLf
    s = s & "<div class=mergebox id=mergePanel hidden>" & vbCrLf
    s = s & "  <div class=mergecols>" & vbCrLf
    s = s & "    <div class=mergecol>" & vbCrLf
    s = s & "      <div class=mh>親をまとめる</div>" & vbCrLf
    s = s & "      <div class=boxes id=atboxes></div>" & vbCrLf
    s = s & "      <div class=mrow><input id=atname placeholder=""まとめ名（任意）"" maxlength=40><button id=atmerge class=mbtn>まとめる</button></div>" & vbCrLf
    s = s & "      <div class=msg id=atmsg></div>" & vbCrLf
    s = s & "      <div class=mtags id=atmerges></div>" & vbCrLf
    s = s & "    </div>" & vbCrLf
    s = s & "    <div class=mergecol>" & vbCrLf
    s = s & "      <div class=mh>子をまとめる</div>" & vbCrLf
    s = s & "      <div class=boxes id=nsboxes></div>" & vbCrLf
    s = s & "      <div class=mrow><input id=nsname placeholder=""まとめ名（任意）"" maxlength=40><button id=nsmerge class=mbtn>まとめる</button></div>" & vbCrLf
    s = s & "      <div class=msg id=nsmsg></div>" & vbCrLf
    s = s & "      <div class=mtags id=nsmerges></div>" & vbCrLf
    s = s & "    </div>" & vbCrLf
    s = s & "  </div>" & vbCrLf
    s = s & "  <div class=mfoot><button id=mergeReset class=clearbtn>まとめを全部解除</button><span class=hint>この操作はこの画面だけの一時表示です。Excel には影響しません。まとめ名は保存されません。</span></div>" & vbCrLf
    s = s & "</div>" & vbCrLf
    s = s & "<div class=mergebox id=presetPanel hidden>" & vbCrLf
    s = s & "  <div class=mh>まとめルール（複数のまとめを1つのルールとして覚える）</div>" & vbCrLf
    s = s & "  <div class=mrow>" & vbCrLf
    s = s & "    <input id=presetName placeholder=""ルール名（例：顧客報告用）"" maxlength=40>" & vbCrLf
    s = s & "    <button id=presetSave class=mbtn>いまのまとめを保存</button>" & vbCrLf
    s = s & "  </div>" & vbCrLf
    s = s & "  <div class=msg id=presetMsg></div>" & vbCrLf
    s = s & "  <div id=presetList class=plist></div>" & vbCrLf
    s = s & "  <div class=mfoot>" & vbCrLf
    s = s & "    <button id=presetClose class=clearbtn>閉じる</button>" & vbCrLf
    s = s & "    <button id=presetExport class=clearbtn>ファイルに書き出す</button>" & vbCrLf
    s = s & "    <label class=clearbtn id=impLabel>ファイルから読み込む<input id=presetImport type=file accept="".json,application/json"" hidden></label>" & vbCrLf
    s = s & "    <span class=hint>ルールはこの端末のブラウザに覚えます。呼び出すと、そのルールのまとめを一括で適用します。Excel には影響しません。</span>" & vbCrLf
    s = s & "  </div>" & vbCrLf
    s = s & "</div>" & vbCrLf
    s = s & "<div class=mergebox id=viewPanel hidden>" & vbCrLf
    s = s & "  <div class=vrow><label class=vlab>絞り込み（子ステータス名）</label><input id=nsFilter placeholder=""キーワード"" maxlength=40><button id=nsFilterClear class=clearbtn>クリア</button></div>" & vbCrLf
    s = s & "  <div class=vrow><label class=vlab>並べ替え（親・子）</label><select id=sortSel><option value=appear>出現順</option><option value=count>多い順</option><option value=name>名前順</option></select></div>" & vbCrLf
    s = s & "  <div class=vrow><label class=vlab>表示</label><button id=darkBtn class=clearbtn aria-pressed=false>ダーク表示</button><button id=cbBtn class=clearbtn aria-pressed=false>色覚に配慮した配色</button></div>" & vbCrLf
    s = s & "  <div class=vrow><label class=vlab>PDFの見出し</label><input id=pdfTitle placeholder=""タイトル"" maxlength=60><input id=pdfDate placeholder=""日付"" maxlength=20><input id=pdfMemo placeholder=""メモ"" maxlength=80></div>" & vbCrLf
    s = s & "  <div class=hint>絞り込み・並べ替え・表示・PDF見出しは、この画面だけの見え方です。Excel には影響しません。</div>" & vbCrLf
    s = s & "</div>" & vbCrLf
    s = s & "<div class=card>" & vbCrLf
    s = s & "  <div class=chart id=chart></div>" & vbCrLf
    s = s & "  <div class=detail id=detail>棒の区画か表のセルをクリックで、その障害IDを表示。棒のラベルか表の見出しで 親を掘り下げ。凡例で 子を絞り込み。棒はドラッグで並べ替えできます。</div>" & vbCrLf
    s = s & "</div>" & vbCrLf
    s = s & "<div class=mergebox id=cellbox hidden>" & vbCrLf
    s = s & "  <div class=cellhead><span id=cellTitle></span><button id=cellClose class=mx>×</button></div>" & vbCrLf
    s = s & "  <div id=cellIds class=cellids></div>" & vbCrLf
    s = s & "</div>" & vbCrLf
    s = s & "<div class=legend id=legend style=""margin:2px 0 8px""></div>" & vbCrLf
    s = s & "<div class=tablewrap><table id=tbl></table></div>" & vbCrLf
    s = s & "</div>" & vbCrLf
    s = s & "<script>" & vbCrLf
    s = s & "const D={{DATA}};" & vbCrLf
    s = s & "const PAL0={""終了"":""#2f8f4e"",""対応計画済"":""#2f5bd0"",""対応中"":""#e08a1e"",""保留・確認待ち"":""#8a5cd0"",""対象外(子)"":""#b0402f"",""（未設定）"":""#9aa7b6""};" & vbCrLf
    s = s & "const EX=[""#4aa3a3"",""#c0762f"",""#6b7280"",""#a83279"",""#3f7a8c"",""#5b8a2f"",""#a8324f"",""#2f6f8c""];" & vbCrLf
    s = s & "const MC=[""#334155"",""#7c3aed"",""#0e7490"",""#9d174d"",""#b45309"",""#15803d""];" & vbCrLf
    s = s & "const CB=[""#e69f00"",""#56b4e9"",""#009e73"",""#d55e00"",""#0072b2"",""#cc79a7"",""#f0e442"",""#999999""];  // 色覚に配慮（Okabe-Ito）" & vbCrLf
    s = s & "const $=function(id){return document.getElementById(id)};" & vbCrLf
    s = s & "let atMerges=[], nsMerges=[];" & vbCrLf
    s = s & "let mode=""count"", atFocus=null, nsFocus=null;" & vbCrLf
    s = s & "let cbSafe=false, nsFilter="""", sortMode=""appear"", splitView=false;" & vbCrLf
    s = s & "let manualAt=null, manualNs=null;   // ドラッグで決めた手動の並び（null=自動）" & vbCrLf
    s = s & "const PXMAX=230;" & vbCrLf
    s = s & "const ATCOL=""#3a5a86"";  // 親の棒の色（対比表示・単色）" & vbCrLf
    s = s & "function esc(s){return String(s).replace(/&/g,""&amp;"").replace(/</g,""&lt;"").replace(/>/g,""&gt;"").replace(/""/g,""&quot;"")}" & vbCrLf
    s = s & "function uniqueName(base,taken){let n=(base||""まとめ"");if(taken.indexOf(n)<0)return n;let i=2;while(taken.indexOf(n+i)>=0)i++;return n+i}" & vbCrLf
    s = s & "function groupOf(x,merges){for(const m of merges){if(m.set.has(x))return m.name}return x}" & vbCrLf
    s = s & "function buildView(){" & vbCrLf
    s = s & "  const aMap=function(a){return groupOf(a,atMerges)}, nMap=function(n){return groupOf(n,nsMerges)};" & vbCrLf
    s = s & "  const at=[]; let ns=[];" & vbCrLf
    s = s & "  D.at.forEach(function(a){const k=aMap(a);if(at.indexOf(k)<0)at.push(k)});" & vbCrLf
    s = s & "  D.ns.forEach(function(n){const k=nMap(n);if(ns.indexOf(k)<0)ns.push(k)});" & vbCrLf
    s = s & "  const srcIds=D.ids||{};" & vbCrLf
    s = s & "  const tbl={},ids={};" & vbCrLf
    s = s & "  at.forEach(function(a){tbl[a]={};ids[a]={};ns.forEach(function(n){tbl[a][n]=0;ids[a][n]=[]})});" & vbCrLf
    s = s & "  D.at.forEach(function(a){D.ns.forEach(function(n){" & vbCrLf
    s = s & "    const da=aMap(a),dn=nMap(n); tbl[da][dn]+=(D.tbl[a][n]||0);" & vbCrLf
    s = s & "    const src=(srcIds[a]&&srcIds[a][n])?srcIds[a][n]:null; if(src&&src.length)ids[da][dn]=ids[da][dn].concat(src);" & vbCrLf
    s = s & "  })});" & vbCrLf
    s = s & "  const nsFull=ns.length;" & vbCrLf
    s = s & "  if(nsFilter){const kw=nsFilter.toLowerCase();const keep=ns.filter(function(n){return n.toLowerCase().indexOf(kw)>=0});if(keep.length)ns=keep;}" & vbCrLf
    s = s & "  const rowtot={};at.forEach(function(a){rowtot[a]=ns.reduce(function(s,n){return s+tbl[a][n]},0)});" & vbCrLf
    s = s & "  if(sortMode===""count""){" & vbCrLf
    s = s & "    const cs={};ns.forEach(function(n){cs[n]=at.reduce(function(s,a){return s+tbl[a][n]},0)});" & vbCrLf
    s = s & "    at.sort(function(x,y){return rowtot[y]-rowtot[x]}); ns.sort(function(x,y){return cs[y]-cs[x]});" & vbCrLf
    s = s & "  }else if(sortMode===""name""){" & vbCrLf
    s = s & "    at.sort(function(x,y){return String(x).localeCompare(String(y),""ja"")}); ns.sort(function(x,y){return String(x).localeCompare(String(y),""ja"")});" & vbCrLf
    s = s & "  }" & vbCrLf
    s = s & "  // ドラッグで決めた手動の並びがあれば、それを最後に当てる（未指定は末尾へ）" & vbCrLf
    s = s & "  if(manualAt){var ia=function(x){var i=manualAt.indexOf(x);return i<0?1e9:i};at.sort(function(x,y){return ia(x)-ia(y)})}" & vbCrLf
    s = s & "  if(manualNs){var ins=function(x){var i=manualNs.indexOf(x);return i<0?1e9:i};ns.sort(function(x,y){return ins(x)-ins(y)})}" & vbCrLf
    s = s & "  const colsum=ns.map(function(n){return at.reduce(function(s,a){return s+tbl[a][n]},0)});" & vbCrLf
    s = s & "  const total=colsum.reduce(function(s,v){return s+v},0);" & vbCrLf
    s = s & "  const PAL={};let ei=0;" & vbCrLf
    s = s & "  ns.forEach(function(n){const mi=nsMerges.findIndex(function(m){return m.name===n});" & vbCrLf
    s = s & "    PAL[n]= mi>=0 ? MC[mi%MC.length] : (cbSafe ? CB[ei++%CB.length] : (PAL0[n]||EX[ei++%EX.length]))});" & vbCrLf
    s = s & "  return {at:at,ns:ns,tbl:tbl,ids:ids,rowtot:rowtot,colsum:colsum,total:total,nsFull:nsFull,PAL:PAL,maxTot:Math.max(1,...at.map(function(a){return rowtot[a]}))};" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "let V=buildView();" & vbCrLf
    s = s & "function renderLegend(){" & vbCrLf
    s = s & "  $(""legend"").innerHTML=V.ns.map(function(n){return '<span class=lg data-ns=""'+esc(n)+'"" aria-pressed=""'+(nsFocus===n)+'""><i style=""background:'+V.PAL[n]+'""></i>'+esc(n)+'</span>'}).join("""");" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "function renderChart(){" & vbCrLf
    s = s & "  const ch=$(""chart"");" & vbCrLf
    s = s & "  if(splitView){ ch.classList.add(""split""); renderSplit(); return; }" & vbCrLf
    s = s & "  ch.classList.remove(""split"");" & vbCrLf
    s = s & "  ch.innerHTML=V.at.map(function(a){" & vbCrLf
    s = s & "    const segs=V.ns.map(function(n){const v=V.tbl[a][n]||0;return v?'<div class=""seg"" data-at=""'+esc(a)+'"" data-ns=""'+esc(n)+'"" data-v=""'+v+'"" style=""background:'+V.PAL[n]+'""></div>':""""}).join("""");" & vbCrLf
    s = s & "    return '<div class=""col"" data-at=""'+esc(a)+'"" data-axis=""at"" data-key=""'+esc(a)+'"" draggable=""true""><div class=""stack"">'+segs+'</div><div class=""xl""><span class=""nm"" title=""'+esc(a)+'"">'+esc(a)+'</span><b>'+V.rowtot[a]+'</b></div></div>';" & vbCrLf
    s = s & "  }).join("""");" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "// 対比表示：左＝親の合計、右＝子の合計を、同じ目盛りで単純な棒に。" & vbCrLf
    s = s & "function renderSplit(){" & vbCrLf
    s = s & "  function bar(kind,key,v,color){" & vbCrLf
    s = s & "    return '<div class=""col"" data-kind=""'+kind+'"" data-axis=""'+kind+'"" data-key=""'+esc(key)+'"" draggable=""true"">'" & vbCrLf
    s = s & "      +'<div class=""stack solid""><div class=""seg"" data-v=""'+v+'"" style=""background:'+color+'""></div></div>'" & vbCrLf
    s = s & "      +'<div class=""xl""><span class=""nm"" title=""'+esc(key)+'"">'+esc(key)+'</span><b>'+v+'</b></div></div>';" & vbCrLf
    s = s & "  }" & vbCrLf
    s = s & "  const left=V.at.map(function(a){return bar(""at"",a,V.rowtot[a],ATCOL)}).join("""");" & vbCrLf
    s = s & "  const right=V.ns.map(function(n,j){return bar(""ns"",n,V.colsum[j],V.PAL[n])}).join("""");" & vbCrLf
    s = s & "  $(""chart"").innerHTML=" & vbCrLf
    s = s & "    '<div class=""splitgroup""><div class=""ghd"">親</div><div class=""gbars"">'+left+'</div></div>'" & vbCrLf
    s = s & "    +'<div class=""split-div""></div>'" & vbCrLf
    s = s & "    +'<div class=""splitgroup""><div class=""ghd"">子</div><div class=""gbars"">'+right+'</div></div>';" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "function renderTable(){" & vbCrLf
    s = s & "  const head='<thead><tr><th class=rowh>親 ＼ 子</th>'+V.ns.map(function(n){return '<th title=""'+esc(n)+'"">'+esc(n)+'</th>'}).join("""")+'<th class=tot>合計</th></tr></thead>';" & vbCrLf
    s = s & "  const body='<tbody>'+V.at.map(function(a){return '<tr data-at=""'+esc(a)+'""><th class=rowh title=""'+esc(a)+'"">'+esc(a)+'</th>'+V.ns.map(function(n){return '<td class=idcell data-ns=""'+esc(n)+'"">'+(V.tbl[a][n]||"""")+'</td>'}).join("""")+'<td class=tot>'+V.rowtot[a]+'</td></tr>'}).join("""")+'</tbody>';" & vbCrLf
    s = s & "  const foot='<tfoot><tr><th class=rowh>合計</th>'+V.colsum.map(function(v){return '<td>'+v+'</td>'}).join("""")+'<td class=tot>'+V.total+'</td></tr></tfoot>';" & vbCrLf
    s = s & "  $(""tbl"").innerHTML=head+body+foot;" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "function draw(){" & vbCrLf
    s = s & "  if(splitView){ drawSplit(); return; }" & vbCrLf
    s = s & "  document.querySelectorAll(""#chart .seg"").forEach(function(s){" & vbCrLf
    s = s & "    const a=s.dataset.at,v=+s.dataset.v;" & vbCrLf
    s = s & "    const h= mode===""pct"" ? v/(V.rowtot[a]||1)*PXMAX : v/V.maxTot*PXMAX;" & vbCrLf
    s = s & "    s.style.height=h+""px"";" & vbCrLf
    s = s & "    s.textContent = h>15 ? (mode===""pct""?Math.round(v/V.rowtot[a]*100)+""%"":v) : """";" & vbCrLf
    s = s & "  });" & vbCrLf
    s = s & "  applyDim();" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "// 対比表示の高さ：親・子を同じ目盛り（両者の最大）でそろえる" & vbCrLf
    s = s & "function drawSplit(){" & vbCrLf
    s = s & "  const segs=Array.prototype.slice.call(document.querySelectorAll(""#chart .seg""));" & vbCrLf
    s = s & "  const max=Math.max(1,...segs.map(function(s){return +s.dataset.v}));" & vbCrLf
    s = s & "  segs.forEach(function(s){const v=+s.dataset.v;const h=v/max*PXMAX;s.style.height=h+""px"";s.textContent=h>14?v:""""});" & vbCrLf
    s = s & "  $(""clear"").hidden=!(atFocus||nsFocus);" & vbCrLf
    s = s & "  detailText();" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "function applyDim(){" & vbCrLf
    s = s & "  if(splitView){ $(""clear"").hidden=!(atFocus||nsFocus); detailText(); return; }" & vbCrLf
    s = s & "  document.querySelectorAll(""#chart .seg"").forEach(function(s){s.classList.toggle(""mute"", !!(nsFocus&&s.dataset.ns!==nsFocus))});" & vbCrLf
    s = s & "  document.querySelectorAll(""#chart .col"").forEach(function(c){c.classList.toggle(""dim"", !!(atFocus&&c.dataset.at!==atFocus));c.classList.toggle(""on"", !!(atFocus&&c.dataset.at===atFocus))});" & vbCrLf
    s = s & "  document.querySelectorAll(""#tbl tbody tr"").forEach(function(tr){tr.classList.toggle(""on"", !!(atFocus&&tr.dataset.at===atFocus))});" & vbCrLf
    s = s & "  const jsel=nsFocus?V.ns.indexOf(nsFocus):-1;" & vbCrLf
    s = s & "  document.querySelectorAll(""#tbl tr"").forEach(function(tr){Array.prototype.forEach.call(tr.children,function(cell,idx){cell.classList.toggle(""oncol"", jsel>=0&&idx===jsel+1)})});" & vbCrLf
    s = s & "  $(""clear"").hidden=!(atFocus||nsFocus);" & vbCrLf
    s = s & "  detailText();" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "function detailText(){" & vbCrLf
    s = s & "  const det=$(""detail"");" & vbCrLf
    s = s & "  if(atFocus&&V.at.indexOf(atFocus)>=0){" & vbCrLf
    s = s & "    const parts=V.ns.filter(function(n){return V.tbl[atFocus][n]&&(!nsFocus||n===nsFocus)}).map(function(n){return '<span class=chip><i style=""background:'+V.PAL[n]+'""></i>'+esc(n)+'：'+V.tbl[atFocus][n]+'件（'+Math.round(V.tbl[atFocus][n]/V.rowtot[atFocus]*100)+'%）</span>'}).join("""");" & vbCrLf
    s = s & "    det.innerHTML='<b>'+esc(atFocus)+'</b>（親・計'+V.rowtot[atFocus]+'件）の内訳：<br>'+parts;" & vbCrLf
    s = s & "  }else if(nsFocus&&V.ns.indexOf(nsFocus)>=0){" & vbCrLf
    s = s & "    const j=V.ns.indexOf(nsFocus);" & vbCrLf
    s = s & "    const parts=V.at.filter(function(a){return V.tbl[a][nsFocus]}).map(function(a){return '<span class=chip><i style=""background:'+V.PAL[nsFocus]+'""></i>'+esc(a)+'：'+V.tbl[a][nsFocus]+'件</span>'}).join("""");" & vbCrLf
    s = s & "    det.innerHTML='子＝<b>'+esc(nsFocus)+'</b>（計'+V.colsum[j]+'件）を持つ親の内訳：<br>'+parts;" & vbCrLf
    s = s & "  }else det.textContent=""棒の区画か表のセルをクリックで、その障害IDを表示。棒のラベルか表の見出しで 親を掘り下げ。凡例で 子を絞り込み。棒はドラッグで並べ替えできます。"";" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "function animate(){" & vbCrLf
    s = s & "  const cols=Array.prototype.slice.call(document.querySelectorAll(""#chart .col""));" & vbCrLf
    s = s & "  cols.forEach(function(c){c.querySelectorAll("".seg"").forEach(function(s){s.style.height=""0px""})});" & vbCrLf
    s = s & "  if(matchMedia(""(prefers-reduced-motion:reduce)"").matches){draw();return}" & vbCrLf
    s = s & "  let i=0;(function step(){ if(i>=cols.length){draw();return}" & vbCrLf
    s = s & "    const c=cols[i++];" & vbCrLf
    s = s & "    c.querySelectorAll("".seg"").forEach(function(s){const a=s.dataset.at,v=+s.dataset.v;const hh=mode===""pct""?v/(V.rowtot[a]||1)*PXMAX:v/V.maxTot*PXMAX;s.style.height=hh+""px"";s.textContent=hh>15?(mode===""pct""?Math.round(v/V.rowtot[a]*100)+""%"":v):""""});" & vbCrLf
    s = s & "    setTimeout(step,140);})();" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "function render(withAnim){" & vbCrLf
    s = s & "  V=buildView();" & vbCrLf
    s = s & "  renderLegend();renderChart();renderTable();" & vbCrLf
    s = s & "  $(""meta"").textContent= nsFilter ? (""（絞り込み中：""+V.total+""件 ／ 子 ""+V.ns.length+""／""+V.nsFull+""列）"") : (""（サンプル ""+V.total+"" 件）"");" & vbCrLf
    s = s & "  if(splitView)drawSplit();" & vbCrLf
    s = s & "  else if(withAnim)animate();else draw();" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "function setAt(a){atFocus=(atFocus===a)?null:a;applyDim()}" & vbCrLf
    s = s & "function setNs(n){nsFocus=(nsFocus===n)?null:n;renderLegend();applyDim()}" & vbCrLf
    s = s & "$(""chart"").addEventListener(""click"",function(e){" & vbCrLf
    s = s & "  if(splitView){const c=e.target.closest("".col"");if(!c)return;if(c.dataset.kind===""at"")setAt(c.dataset.key);else setNs(c.dataset.key);return}" & vbCrLf
    s = s & "  const s=e.target.closest("".seg"");if(s){showCellIds(s.dataset.at,s.dataset.ns);return}" & vbCrLf
    s = s & "  const c=e.target.closest("".col"");if(c)setAt(c.dataset.at);" & vbCrLf
    s = s & "});" & vbCrLf
    s = s & "$(""chart"").addEventListener(""mousemove"",function(e){const tip=$(""tip"");" & vbCrLf
    s = s & "  if(splitView){const c=e.target.closest("".col"");if(!c){tip.style.opacity=0;return}const kind=c.dataset.kind===""at""?""親"":""子"";const v=c.querySelector("".seg"").dataset.v;tip.innerHTML=esc(c.dataset.key)+""<br><b>""+v+""件</b>（""+kind+""の合計）"";tip.style.left=(e.clientX+12)+""px"";tip.style.top=(e.clientY+12)+""px"";tip.style.opacity=1;return}" & vbCrLf
    s = s & "  const s=e.target.closest("".seg"");if(!s){tip.style.opacity=0;return}const a=s.dataset.at,n=s.dataset.ns,v=+s.dataset.v;tip.innerHTML=esc(a)+"" × ""+esc(n)+""<br><b>""+v+""件</b>（親内 ""+Math.round(v/V.rowtot[a]*100)+""%）"";tip.style.left=(e.clientX+12)+""px"";tip.style.top=(e.clientY+12)+""px"";tip.style.opacity=1});" & vbCrLf
    s = s & "$(""splitBtn"").addEventListener(""click"",function(){splitView=!splitView;this.setAttribute(""aria-pressed"",String(splitView));render(false)});" & vbCrLf
    s = s & "// ドラッグ&ドロップで棒を並べ替え（同じ側どうしでのみ入れ替え）" & vbCrLf
    s = s & "var dragKey=null, dragAxis=null;" & vbCrLf
    s = s & "$(""chart"").addEventListener(""dragstart"",function(e){var c=e.target.closest("".col"");if(!c)return;dragKey=c.dataset.key;dragAxis=c.dataset.axis;c.classList.add(""dragging"");try{e.dataTransfer.effectAllowed=""move"";e.dataTransfer.setData(""text/plain"",dragKey)}catch(_){}} );" & vbCrLf
    s = s & "$(""chart"").addEventListener(""dragend"",function(){document.querySelectorAll(""#chart .col.dragging,#chart .col.dropok"").forEach(function(x){x.classList.remove(""dragging"");x.classList.remove(""dropok"")})});" & vbCrLf
    s = s & "$(""chart"").addEventListener(""dragover"",function(e){var c=e.target.closest("".col"");if(!c||c.dataset.axis!==dragAxis)return;e.preventDefault();try{e.dataTransfer.dropEffect=""move""}catch(_){}document.querySelectorAll(""#chart .col.dropok"").forEach(function(x){x.classList.remove(""dropok"")});if(c.dataset.key!==dragKey)c.classList.add(""dropok"")});" & vbCrLf
    s = s & "$(""chart"").addEventListener(""drop"",function(e){var c=e.target.closest("".col"");if(!c||c.dataset.axis!==dragAxis||dragKey==null)return;e.preventDefault();var tgt=c.dataset.key;if(tgt===dragKey)return;" & vbCrLf
    s = s & "  var cur=(dragAxis===""at"")?V.at.slice():V.ns.slice();" & vbCrLf
    s = s & "  var from=cur.indexOf(dragKey); if(from>=0)cur.splice(from,1);" & vbCrLf
    s = s & "  var to=cur.indexOf(tgt); if(to<0)to=cur.length; cur.splice(to,0,dragKey);" & vbCrLf
    s = s & "  if(dragAxis===""at"")manualAt=cur; else manualNs=cur;" & vbCrLf
    s = s & "  dragKey=null;dragAxis=null; render(false);" & vbCrLf
    s = s & "});" & vbCrLf
    s = s & "$(""chart"").addEventListener(""mouseleave"",function(){$(""tip"").style.opacity=0});" & vbCrLf
    s = s & "$(""legend"").addEventListener(""click"",function(e){const l=e.target.closest("".lg"");if(l)setNs(l.dataset.ns)});" & vbCrLf
    s = s & "$(""tbl"").addEventListener(""click"",function(e){" & vbCrLf
    s = s & "  const td=e.target.closest(""td.idcell"");" & vbCrLf
    s = s & "  if(td){const tr=td.closest(""tr"");showCellIds(tr.dataset.at,td.dataset.ns);return}" & vbCrLf
    s = s & "  const tr=e.target.closest(""tbody tr""); if(tr&&e.target.closest(""th.rowh""))setAt(tr.dataset.at);" & vbCrLf
    s = s & "});" & vbCrLf
    s = s & "$(""clear"").addEventListener(""click"",function(){atFocus=null;nsFocus=null;renderLegend();applyDim()});" & vbCrLf
    s = s & "document.querySelectorAll("".seg-ctl button"").forEach(function(b){b.addEventListener(""click"",function(){mode=b.dataset.m;document.querySelectorAll("".seg-ctl button"").forEach(function(x){x.setAttribute(""aria-pressed"",x===b)});draw()})});" & vbCrLf
    s = s & "// PDF出力（A4）。ブラウザの印刷でPDFに保存。まとめ後の状態もそのまま出る。" & vbCrLf
    s = s & "$(""pdfBtn"").addEventListener(""click"",function(){window.print()});" & vbCrLf
    s = s & "window.addEventListener(""beforeprint"",function(){draw()});" & vbCrLf
    s = s & "// ---- まとめ（マージ）----" & vbCrLf
    s = s & "function takenSet(merges){const s=new Set();merges.forEach(function(m){m.set.forEach(function(x){s.add(x)})});return s}" & vbCrLf
    s = s & "function renderBuilder(){" & vbCrLf
    s = s & "  const at_t=takenSet(atMerges), ns_t=takenSet(nsMerges);" & vbCrLf
    s = s & "  const ab=D.at.filter(function(a){return !at_t.has(a)});" & vbCrLf
    s = s & "  const nb=D.ns.filter(function(n){return !ns_t.has(n)});" & vbCrLf
    s = s & "  $(""atboxes"").innerHTML= ab.length? ab.map(function(a){return '<label class=ck><input type=checkbox value=""'+esc(a)+'"">'+esc(a)+'</label>'}).join("""") : '<span class=muted>まとめられる項目がありません</span>';" & vbCrLf
    s = s & "  $(""nsboxes"").innerHTML= nb.length? nb.map(function(n){return '<label class=ck><input type=checkbox value=""'+esc(n)+'""><i style=""background:'+(PAL0[n]||""#8a97a8"")+'""></i>'+esc(n)+'</label>'}).join("""") : '<span class=muted>まとめられる項目がありません</span>';" & vbCrLf
    s = s & "  $(""atmerges"").innerHTML=atMerges.map(function(m,i){return '<span class=mtag>'+esc(m.name)+'<button class=mx data-side=at data-i=""'+i+'"">×</button></span>'}).join("""");" & vbCrLf
    s = s & "  $(""nsmerges"").innerHTML=nsMerges.map(function(m,i){return '<span class=mtag>'+esc(m.name)+'<button class=mx data-side=ns data-i=""'+i+'"">×</button></span>'}).join("""");" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "function doMerge(side){" & vbCrLf
    s = s & "  const merges= side===""at""?atMerges:nsMerges;" & vbCrLf
    s = s & "  const picked=Array.prototype.slice.call(document.querySelectorAll(""#""+side+""boxes input:checked"")).map(function(x){return x.value});" & vbCrLf
    s = s & "  const msg=$(side+""msg"");" & vbCrLf
    s = s & "  if(picked.length<2){msg.textContent=""2つ以上えらんでください。"";return}" & vbCrLf
    s = s & "  msg.textContent="""";" & vbCrLf
    s = s & "  const orig= side===""at""?D.at:D.ns;" & vbCrLf
    s = s & "  // まとめに吸収される項目（今回選んだもの＋既存のまとめのメンバー）は衝突相手にしない。" & vbCrLf
    s = s & "  // まとめ後に別個で残る名前（＝未まとめの項目＋既存のまとめ名）だけと重複判定する。" & vbCrLf
    s = s & "  const pickedSet=new Set(picked), inMerges=takenSet(merges);" & vbCrLf
    s = s & "  const taken=orig.filter(function(x){return !pickedSet.has(x) && !inMerges.has(x)}).concat(merges.map(function(m){return m.name}));" & vbCrLf
    s = s & "  const name=uniqueName(($(side+""name"").value||"""").trim()||""まとめ"", taken);" & vbCrLf
    s = s & "  merges.push({name:name,set:new Set(picked)});" & vbCrLf
    s = s & "  $(side+""name"").value="""";" & vbCrLf
    s = s & "  atFocus=null;nsFocus=null;" & vbCrLf
    s = s & "  renderBuilder();render(false);" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "function removeMerge(side,i){(side===""at""?atMerges:nsMerges).splice(i,1);atFocus=null;nsFocus=null;renderBuilder();render(false)}" & vbCrLf
    s = s & "$(""mergeToggle"").addEventListener(""click"",function(){const p=$(""mergePanel"");p.hidden=!p.hidden;$(""mergeToggle"").setAttribute(""aria-pressed"",!p.hidden);if(!p.hidden)renderBuilder()});" & vbCrLf
    s = s & "$(""atmerge"").addEventListener(""click"",function(){doMerge(""at"")});" & vbCrLf
    s = s & "$(""nsmerge"").addEventListener(""click"",function(){doMerge(""ns"")});" & vbCrLf
    s = s & "$(""mergeReset"").addEventListener(""click"",function(){atMerges=[];nsMerges=[];atFocus=null;nsFocus=null;renderBuilder();render(false)});" & vbCrLf
    s = s & "$(""mergePanel"").addEventListener(""click"",function(e){const b=e.target.closest("".mx"");if(b)removeMerge(b.dataset.side,+b.dataset.i)});" & vbCrLf
    s = s & "// ---- まとめルールの保存・呼び出し（複数のまとめを1つのルールとして覚える）----" & vbCrLf
    s = s & "var PKEY=""ns_dash_presets_v1"";" & vbCrLf
    s = s & "function loadPresets(){try{return JSON.parse(localStorage.getItem(PKEY)||""[]"")||[]}catch(e){return []}}" & vbCrLf
    s = s & "function savePresets(arr){try{localStorage.setItem(PKEY,JSON.stringify(arr));return true}catch(e){return false}}" & vbCrLf
    s = s & "function currentRule(){return {at:atMerges.map(function(m){return {name:m.name,members:Array.from(m.set)}}),ns:nsMerges.map(function(m){return {name:m.name,members:Array.from(m.set)}})}}" & vbCrLf
    s = s & "function applyRule(rule){" & vbCrLf
    s = s & "  var mk=function(defs,orig){return (defs||[]).map(function(d){return {name:d.name,set:new Set((d.members||[]).filter(function(x){return orig.indexOf(x)>=0}))}}).filter(function(m){return m.set.size>=2})};" & vbCrLf
    s = s & "  atMerges=mk(rule.at,D.at); nsMerges=mk(rule.ns,D.ns);" & vbCrLf
    s = s & "  atFocus=null;nsFocus=null; renderBuilder(); render(false);" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "function renderPresetList(){" & vbCrLf
    s = s & "  var ps=loadPresets();" & vbCrLf
    s = s & "  $(""presetList"").innerHTML= ps.length? ps.map(function(p,i){return '<div class=prow><button class=applybtn data-i=""'+i+'"">'+esc(p.name)+'</button><span class=pcount>親'+((p.at||[]).length)+'・子'+((p.ns||[]).length)+'</span><button class=mx data-del=""'+i+'"">×</button></div>'}).join("""") : '<div class=muted>保存したルールはまだありません。まとめてから「いまのまとめを保存」を押してください。</div>';" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "function openPreset(){$(""presetPanel"").hidden=false;$(""presetMsg"").className=""msg"";$(""presetMsg"").textContent="""";renderPresetList()}" & vbCrLf
    s = s & "function doSavePreset(){" & vbCrLf
    s = s & "  var msg=$(""presetMsg""); msg.className=""msg"";" & vbCrLf
    s = s & "  if(!atMerges.length && !nsMerges.length){msg.textContent=""先にステータスをまとめてください。"";return}" & vbCrLf
    s = s & "  var name=($(""presetName"").value||"""").trim();" & vbCrLf
    s = s & "  if(!name){msg.textContent=""ルール名を入れてください。"";$(""presetName"").focus();return}" & vbCrLf
    s = s & "  var ps=loadPresets(), rule=currentRule(); rule.name=name;" & vbCrLf
    s = s & "  var i=-1; ps.forEach(function(p,k){if(p.name===name)i=k});" & vbCrLf
    s = s & "  if(i>=0)ps[i]=rule; else ps.push(rule);" & vbCrLf
    s = s & "  var ok=savePresets(ps);" & vbCrLf
    s = s & "  $(""presetName"").value=""""; renderPresetList();" & vbCrLf
    s = s & "  msg.className= ok? ""msg ok"":""msg"";" & vbCrLf
    s = s & "  msg.textContent= ok? ('「'+name+'」を保存しました。呼び出しボタンで一括適用できます。') : 'この端末に保存できませんでした。「ファイルに書き出す」で保存してください。';" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "function exportPresets(){" & vbCrLf
    s = s & "  var ps=loadPresets();" & vbCrLf
    s = s & "  if(!ps.length){$(""presetMsg"").textContent=""書き出すルールがありません。"";return}" & vbCrLf
    s = s & "  var a=document.createElement(""a"");" & vbCrLf
    s = s & "  a.href=URL.createObjectURL(new Blob([JSON.stringify(ps,null,1)],{type:""application/json""}));" & vbCrLf
    s = s & "  a.download=""まとめルール.json""; document.body.appendChild(a); a.click(); a.remove();" & vbCrLf
    s = s & "  setTimeout(function(){URL.revokeObjectURL(a.href)},1000);" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "function importPresets(file){" & vbCrLf
    s = s & "  var rd=new FileReader();" & vbCrLf
    s = s & "  rd.onload=function(){try{" & vbCrLf
    s = s & "    var arr=JSON.parse(rd.result);" & vbCrLf
    s = s & "    if(!Array.isArray(arr)){$(""presetMsg"").textContent=""ファイルの形式が違います。"";return}" & vbCrLf
    s = s & "    var cur=loadPresets();" & vbCrLf
    s = s & "    arr.forEach(function(p){if(p&&p.name){var i=-1;cur.forEach(function(x,k){if(x.name===p.name)i=k});if(i>=0)cur[i]=p;else cur.push(p)}});" & vbCrLf
    s = s & "    savePresets(cur); renderPresetList(); $(""presetMsg"").className=""msg ok""; $(""presetMsg"").textContent=""読み込みました（""+arr.length+""件）。"";" & vbCrLf
    s = s & "  }catch(e){$(""presetMsg"").textContent=""読み込めませんでした。""}};" & vbCrLf
    s = s & "  rd.readAsText(file);" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "$(""saveBtn"").addEventListener(""click"",function(){openPreset();$(""presetName"").focus()});" & vbCrLf
    s = s & "$(""loadBtn"").addEventListener(""click"",function(){openPreset()});" & vbCrLf
    s = s & "$(""presetSave"").addEventListener(""click"",doSavePreset);" & vbCrLf
    s = s & "$(""presetName"").addEventListener(""keydown"",function(e){if(e.key===""Enter"")doSavePreset()});" & vbCrLf
    s = s & "$(""presetClose"").addEventListener(""click"",function(){$(""presetPanel"").hidden=true});" & vbCrLf
    s = s & "$(""presetExport"").addEventListener(""click"",exportPresets);" & vbCrLf
    s = s & "$(""presetImport"").addEventListener(""change"",function(e){if(e.target.files&&e.target.files[0]){importPresets(e.target.files[0]);e.target.value=""""}});" & vbCrLf
    s = s & "$(""presetPanel"").addEventListener(""click"",function(e){" & vbCrLf
    s = s & "  var ap=e.target.closest("".applybtn""); if(ap){var p=loadPresets()[+ap.dataset.i];if(p){applyRule(p);$(""presetMsg"").className=""msg ok"";$(""presetMsg"").textContent='「'+esc(p.name)+'」を適用しました。'}return}" & vbCrLf
    s = s & "  var del=e.target.closest(""[data-del]""); if(del){var ps=loadPresets();ps.splice(+del.dataset.del,1);savePresets(ps);renderPresetList();return}" & vbCrLf
    s = s & "});" & vbCrLf
    s = s & "// ---- セルの障害ID表示 ----" & vbCrLf
    s = s & "function showCellIds(a,n){" & vbCrLf
    s = s & "  var arr=(V.ids[a]&&V.ids[a][n])?V.ids[a][n]:[];" & vbCrLf
    s = s & "  $(""cellTitle"").innerHTML=esc(a)+' × '+esc(n)+'（'+arr.length+'件）';" & vbCrLf
    s = s & "  $(""cellIds"").innerHTML= arr.length? arr.map(function(id){return '<span class=idchip>'+esc(id)+'</span>'}).join("""") : '<span class=muted>該当する障害IDはありません</span>';" & vbCrLf
    s = s & "  $(""cellbox"").hidden=false; $(""cellbox"").scrollIntoView({block:""nearest""});" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "$(""cellClose"").addEventListener(""click"",function(){$(""cellbox"").hidden=true});" & vbCrLf
    s = s & "// ---- トースト ----" & vbCrLf
    s = s & "var toastT;" & vbCrLf
    s = s & "function toast(m){var t=$(""toast"");t.textContent=m;t.style.opacity=""1"";clearTimeout(toastT);toastT=setTimeout(function(){t.style.opacity=""0""},1800)}" & vbCrLf
    s = s & "// ---- CSV（いまの表・まとめ後の状態）----" & vbCrLf
    s = s & "function csvOfView(){" & vbCrLf
    s = s & "  var rows=[[""親＼子""].concat(V.ns).concat([""合計""])];" & vbCrLf
    s = s & "  V.at.forEach(function(a){rows.push([a].concat(V.ns.map(function(n){return V.tbl[a][n]||0})).concat([V.rowtot[a]]))});" & vbCrLf
    s = s & "  rows.push([""合計""].concat(V.colsum).concat([V.total]));" & vbCrLf
    s = s & "  return rows.map(function(r){return r.map(function(c){var s=String(c);return /["",\n]/.test(s)?'""'+s.replace(/""/g,'""""')+'""':s}).join("","")}).join(""\r\n"");" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "function downloadCsv(csv){var a=document.createElement(""a"");a.href=URL.createObjectURL(new Blob([""﻿""+csv],{type:""text/csv""}));a.download=""障害ステータス集計.csv"";document.body.appendChild(a);a.click();a.remove();setTimeout(function(){URL.revokeObjectURL(a.href)},1000)}" & vbCrLf
    s = s & "function fallbackCopy(text,done){try{var ta=document.createElement(""textarea"");ta.value=text;ta.style.position=""fixed"";ta.style.opacity=""0"";document.body.appendChild(ta);ta.focus();ta.select();var ok=document.execCommand(""copy"");ta.remove();done(ok)}catch(e){done(false)}}" & vbCrLf
    s = s & "$(""csvBtn"").addEventListener(""click"",function(){" & vbCrLf
    s = s & "  var csv=csvOfView();" & vbCrLf
    s = s & "  var done=function(ok){ if(ok){toast(""表をコピーしました（CSV）"")} else {toast(""コピーできないので、CSVを保存します"");downloadCsv(csv)} };" & vbCrLf
    s = s & "  if(navigator.clipboard&&navigator.clipboard.writeText){navigator.clipboard.writeText(csv).then(function(){done(true)},function(){fallbackCopy(csv,done)})}" & vbCrLf
    s = s & "  else fallbackCopy(csv,done);" & vbCrLf
    s = s & "});" & vbCrLf
    s = s & "// ---- 表示設定（絞り込み・並べ替え・ダーク・色覚配慮・PDF見出し）----" & vbCrLf
    s = s & "$(""viewToggle"").addEventListener(""click"",function(){var p=$(""viewPanel"");p.hidden=!p.hidden;this.setAttribute(""aria-pressed"",String(!p.hidden))});" & vbCrLf
    s = s & "$(""nsFilter"").addEventListener(""input"",function(){nsFilter=this.value.trim();atFocus=null;nsFocus=null;$(""cellbox"").hidden=true;render(false)});" & vbCrLf
    s = s & "$(""nsFilterClear"").addEventListener(""click"",function(){nsFilter="""";$(""nsFilter"").value="""";render(false)});" & vbCrLf
    s = s & "$(""sortSel"").addEventListener(""change"",function(){sortMode=this.value;manualAt=null;manualNs=null;render(false)});" & vbCrLf
    s = s & "$(""darkBtn"").addEventListener(""click"",function(){var on=document.documentElement.getAttribute(""data-theme"")!==""dark"";document.documentElement.setAttribute(""data-theme"",on?""dark"":""light"");this.setAttribute(""aria-pressed"",String(on))});" & vbCrLf
    s = s & "$(""cbBtn"").addEventListener(""click"",function(){cbSafe=!cbSafe;this.setAttribute(""aria-pressed"",String(cbSafe));render(false)});" & vbCrLf
    s = s & "try{$(""pdfDate"").value=new Date().toLocaleDateString(""ja-JP"")}catch(e){}" & vbCrLf
    s = s & "window.addEventListener(""beforeprint"",function(){" & vbCrLf
    s = s & "  var t=($(""pdfTitle"").value||"""").trim(),d=($(""pdfDate"").value||"""").trim(),m=($(""pdfMemo"").value||"""").trim();" & vbCrLf
    s = s & "  $(""printHead"").innerHTML=(t?'<div class=ph-title>'+esc(t)+'</div>':'')+((d||m)?'<div class=ph-sub>'+esc(d)+(d&&m?'　':'')+esc(m)+'</div>':'');" & vbCrLf
    s = s & "});" & vbCrLf
    s = s & "render(true);" & vbCrLf
    s = s & "</script></body></html>" & vbCrLf
    s = s & "" & vbCrLf
    HTMLテンプレ = s
End Function
