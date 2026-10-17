param([string]$OutPath = (Join-Path $PSScriptRoot "sample.xls"))

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -lt 3) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }

  $ws = $wb.Worksheets.Item(1)
  $ws.Name = "一覧"
  $headers = @("機種", "図番", "向けさき", "作番", "本体価格", "制御盤", "オプション", "オプション価格")
  for ($i = 0; $i -lt $headers.Count; $i++) { $ws.Cells.Item(3, $i + 1).Value2 = $headers[$i] }
  foreach ($r in 4, 5, 7, 8) {
    $ws.Cells.Item($r, 1).Value2 = "A"
    $ws.Cells.Item($r, 2).Value2 = "C"
    $ws.Cells.Item($r, 3).Value2 = "V"
    $ws.Cells.Item($r, 4).Value2 = "BB"
    $ws.Cells.Item($r, 5).Value2 = 80000
    $ws.Cells.Item($r, 7).Value2 = "1,3,5"
    $ws.Cells.Item($r, 8).Value2 = 9587
  }
  $ws.Range("E14").Formula = "=SUM(E4:E13)"
  $ws.Range("H14").Formula = "=SUM(H4:H13)"
  $ws.Range("A3:H14").Borders.LineStyle = 1
  $ws.Cells.Item(40, 1).Value2 = "下の方の行"
  $ws.Cells.Item(40, 5).Value2 = 1

  $ws2 = $wb.Worksheets.Item(2)
  $ws2.Name = "マスタ"
  $ws2.Range("A1:A300").Formula = "=ROW()"
  $ws2.Range("B1:B300").Formula = '="品目"&ROW()'
  $ws2.Range("C1:C300").Formula = "=ROW()*100"
  $ws2.Range("D1").Formula = "=SUM(C1:C300)"

  $ws3 = $wb.Worksheets.Item(3)
  $ws3.Name = "メモ"
  $ws3.Range("A1").Value2 = "参照"
  $ws3.Range("B1").Formula = "=一覧!E14"

  if (Test-Path $OutPath) { Remove-Item $OutPath }
  $wb.SaveAs($OutPath, 56)
  $wb.Close($false)
  "saved $OutPath"
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
