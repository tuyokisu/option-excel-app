param(
  [string]$OutPath = (Join-Path $PSScriptRoot "該当オプション_値段.xlsx")
)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $ws.Name = "仕様"

  $ws.Cells.Item(1, 1).Formula = "該当するオプション（値段つき見本）"
  $ws.Cells.Item(2, 1).Formula = "左＝使う側（番号・品名・値段・○）"
  $ws.Cells.Item(2, 6).Formula = "右＝使わない側"

  $ws.Cells.Item(3, 1).Formula = "番号"
  $ws.Cells.Item(3, 2).Formula = "品名"
  $ws.Cells.Item(3, 3).Formula = "値段"
  $ws.Cells.Item(3, 4).Formula = "選択"
  $ws.Cells.Item(3, 6).Formula = "番号"
  $ws.Cells.Item(3, 7).Formula = "品名"
  $ws.Cells.Item(3, 8).Formula = "値段"
  $ws.Cells.Item(3, 9).Formula = "選択"

  $left = @(
    @{ no = 1; name = "サンルーフ"; price = 50000; mark = $true },
    @{ no = 2; name = "ナビ"; price = 80000; mark = $false },
    @{ no = 3; name = "エアロ"; price = 120000; mark = $true },
    @{ no = 4; name = "アルミホイール"; price = 90000; mark = $false },
    @{ no = 5; name = "バックカメラ"; price = 35000; mark = $true }
  )
  $right = @(
    @{ no = 9; name = "フロアマット"; price = 999; mark = $true },
    @{ no = 10; name = "ドラレコ"; price = 888; mark = $true },
    @{ no = 11; name = "シートカバー"; price = 777; mark = $false },
    @{ no = 12; name = "キャリア"; price = 666; mark = $true },
    @{ no = 13; name = "ドアバイザー"; price = 555; mark = $false }
  )

  for ($i = 0; $i -lt $left.Count; $i++) {
    $nameRow = 4 + ($i * 2)
    $markRow = $nameRow + 1
    $ws.Cells.Item($nameRow, 1).Formula = [string]$left[$i].no
    $ws.Cells.Item($nameRow, 2).Formula = [string]$left[$i].name
    $ws.Cells.Item($nameRow, 3).Formula = [string]$left[$i].price
    if ($left[$i].mark) {
      $ws.Cells.Item($markRow, 4).Formula = "○"
    }
    $ws.Cells.Item($nameRow, 6).Formula = [string]$right[$i].no
    $ws.Cells.Item($nameRow, 7).Formula = [string]$right[$i].name
    $ws.Cells.Item($nameRow, 8).Formula = [string]$right[$i].price
    if ($right[$i].mark) {
      $ws.Cells.Item($markRow, 9).Formula = "○"
    }
  }

  1..9 | ForEach-Object {
    if ($_ -eq 5) { $ws.Columns.Item($_).ColumnWidth = 3 }
    elseif ($_ -in 2, 7) { $ws.Columns.Item($_).ColumnWidth = 16 }
    elseif ($_ -in 3, 8) { $ws.Columns.Item($_).ColumnWidth = 10 }
    else { $ws.Columns.Item($_).ColumnWidth = 8 }
  }
  $ws.Range("A1:D3").Font.Bold = $true
  $ws.Range("F2:I3").Font.Bold = $true
  $last = 3 + ($left.Count * 2)
  $ws.Range("A3:D$last").Borders.LineStyle = 1
  $ws.Range("F3:I$last").Borders.LineStyle = 1
  $ws.Range("D4:D$last").HorizontalAlignment = -4108
  $ws.Range("I4:I$last").HorizontalAlignment = -4108
  $ws.Cells.Item($last + 2, 1).Formula = "左の ○ は 1,3,5。値段は 50000 / 120000 / 35000。右側は読みません。"

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
