param(
  [string]$OutPath = (Join-Path $PSScriptRoot "該当オプション.xlsx")
)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $ws.Name = "仕様"

  $ws.Cells.Item(1, 1).Formula = "該当するオプション（見本）"
  $ws.Cells.Item(2, 1).Formula = "左＝使う側（番号・品名・○）"
  $ws.Cells.Item(2, 5).Formula = "右＝使わない側"

  $ws.Cells.Item(3, 1).Formula = "番号"
  $ws.Cells.Item(3, 2).Formula = "品名"
  $ws.Cells.Item(3, 3).Formula = "選択"
  $ws.Cells.Item(3, 5).Formula = "番号"
  $ws.Cells.Item(3, 6).Formula = "品名"
  $ws.Cells.Item(3, 7).Formula = "選択"

  $left = @(
    @{ no = 1; name = "サンルーフ"; mark = $true },
    @{ no = 2; name = "ナビ"; mark = $false },
    @{ no = 3; name = "エアロ"; mark = $true },
    @{ no = 4; name = "アルミホイール"; mark = $false },
    @{ no = 5; name = "バックカメラ"; mark = $true }
  )
  $right = @(
    @{ no = 9; name = "フロアマット"; mark = $true },
    @{ no = 10; name = "ドラレコ"; mark = $true },
    @{ no = 11; name = "シートカバー"; mark = $false },
    @{ no = 12; name = "キャリア"; mark = $true },
    @{ no = 13; name = "ドアバイザー"; mark = $false }
  )

  for ($i = 0; $i -lt $left.Count; $i++) {
    $nameRow = 4 + ($i * 2)
    $markRow = $nameRow + 1
    $ws.Cells.Item($nameRow, 1).Formula = [string]$left[$i].no
    $ws.Cells.Item($nameRow, 2).Formula = [string]$left[$i].name
    if ($left[$i].mark) {
      $ws.Cells.Item($markRow, 3).Formula = "○"
    }
    $ws.Cells.Item($nameRow, 5).Formula = [string]$right[$i].no
    $ws.Cells.Item($nameRow, 6).Formula = [string]$right[$i].name
    if ($right[$i].mark) {
      $ws.Cells.Item($markRow, 7).Formula = "○"
    }
  }

  $ws.Columns.Item(1).ColumnWidth = 8
  $ws.Columns.Item(2).ColumnWidth = 18
  $ws.Columns.Item(3).ColumnWidth = 8
  $ws.Columns.Item(4).ColumnWidth = 3
  $ws.Columns.Item(5).ColumnWidth = 8
  $ws.Columns.Item(6).ColumnWidth = 16
  $ws.Columns.Item(7).ColumnWidth = 8
  $ws.Range("A1:C3").Font.Bold = $true
  $ws.Range("E2:G3").Font.Bold = $true
  $last = 3 + ($left.Count * 2)
  $ws.Range("A3:C$last").Borders.LineStyle = 1
  $ws.Range("E3:G$last").Borders.LineStyle = 1
  $ws.Range("C4:C$last").HorizontalAlignment = -4108
  $ws.Range("G4:G$last").HorizontalAlignment = -4108

  $ws.Cells.Item($last + 2, 1).Formula = "この見本では、左の ○ は 1,3,5 です。右側の ○ は読みません。"

  $dir = Split-Path -Parent $OutPath
  if (-not (Test-Path $dir)) {
    New-Item -ItemType Directory -Path $dir | Out-Null
  }
  if (Test-Path $OutPath) {
    Remove-Item -Force $OutPath
  }
  $wb.SaveAs($OutPath, 51)
  $wb.Close($false)
} finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
