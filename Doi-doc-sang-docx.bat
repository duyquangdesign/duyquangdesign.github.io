<# : batch portion
@echo off
title Doi .doc sang .docx - Tro ly Bao Giang
echo Dang mo cong cu doi file .doc sang .docx, vui long doi...
set "TLBG_SELF=%~f0"
powershell -NoProfile -ExecutionPolicy Bypass -STA -Command "$p=$env:TLBG_SELF; iex ([IO.File]::ReadAllText($p, [Text.Encoding]::UTF8))"
exit /b
#>
# ===================== PowerShell =====================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
$title = 'Trợ lý Báo Giảng – Đổi .doc sang .docx'
function Msg($text, $icon = 'Information') { [void][System.Windows.Forms.MessageBox]::Show($text, $title, 'OK', $icon) }

# 1. Chọn thư mục (mặc định: thư mục đang chứa công cụ)
$here = Split-Path -Parent $env:TLBG_SELF
$dlg = New-Object System.Windows.Forms.FolderBrowserDialog
$dlg.Description = 'Chọn thư mục chứa giáo án .doc (công cụ sẽ đổi cả các thư mục con bên trong)'
$dlg.SelectedPath = $here
$dlg.ShowNewFolderButton = $false
if ($dlg.ShowDialog() -ne 'OK') { exit }
$root = $dlg.SelectedPath

# 2. Tìm file .doc chưa có bản .docx
$files = @(Get-ChildItem -LiteralPath $root -Recurse -File -ErrorAction SilentlyContinue |
  Where-Object { $_.Extension -ieq '.doc' -and -not $_.Name.StartsWith('~$') -and
    -not (Test-Path -LiteralPath ([IO.Path]::ChangeExtension($_.FullName, '.docx'))) })
if ($files.Count -eq 0) { Msg "Không có file .doc nào cần đổi trong:`n$root`n`n(Các file .doc đã có bản .docx cùng tên sẽ được bỏ qua.)"; exit }

# 3. Cửa sổ tiến độ
$form = New-Object System.Windows.Forms.Form
$form.Text = $title; $form.Width = 520; $form.Height = 150; $form.StartPosition = 'CenterScreen'; $form.FormBorderStyle = 'FixedDialog'; $form.MaximizeBox = $false; $form.TopMost = $true
$lbl = New-Object System.Windows.Forms.Label; $lbl.Left = 15; $lbl.Top = 15; $lbl.Width = 480; $lbl.Height = 40
$bar = New-Object System.Windows.Forms.ProgressBar; $bar.Left = 15; $bar.Top = 60; $bar.Width = 475; $bar.Maximum = $files.Count
$form.Controls.Add($lbl); $form.Controls.Add($bar); $form.Show(); [System.Windows.Forms.Application]::DoEvents()

$ok = 0; $fail = @(); $engine = ''
# 4a. Dùng Microsoft Word nếu có (giữ nguyên định dạng, hình ảnh)
$word = $null
try { $word = New-Object -ComObject Word.Application; $engine = 'Word' } catch { $word = $null }
if ($word) {
  try { $word.Visible = $false; $word.DisplayAlerts = 0; $word.AutomationSecurity = 3 } catch {}
  $i = 0
  foreach ($f in $files) {
    $i++; $lbl.Text = "Đang đổi ($i/$($files.Count)): $($f.Name)"; $bar.Value = $i - 1; [System.Windows.Forms.Application]::DoEvents()
    $target = [IO.Path]::ChangeExtension($f.FullName, '.docx'); $doc = $null
    try {
      $doc = $word.Documents.Open($f.FullName, $false, $true, $false)
      try { $doc.SaveAs2($target, 16) } catch { $doc.SaveAs($target, 16) }
      $ok++
    } catch { $fail += $f.FullName }
    finally { if ($doc) { try { $doc.Close($false) } catch {} ; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc) } catch {} } }
  }
  try { $word.Quit() } catch {}; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch {}
} else {
  # 4b. Không có Word: thử LibreOffice
  $so = @("$env:ProgramFiles\LibreOffice\program\soffice.exe", "${env:ProgramFiles(x86)}\LibreOffice\program\soffice.exe") | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
  if (-not $so) {
    $form.Close()
    Msg "Máy chưa cài Microsoft Word hoặc LibreOffice nên chưa đổi được.`n`nCách khắc phục: cài LibreOffice miễn phí tại libreoffice.org rồi chạy lại công cụ này, hoặc nhắn Zalo 0382156801 để được hỗ trợ." 'Warning'
    exit
  }
  $engine = 'LibreOffice'; $i = 0
  foreach ($f in $files) {
    $i++; $lbl.Text = "Đang đổi ($i/$($files.Count)): $($f.Name)"; $bar.Value = $i - 1; [System.Windows.Forms.Application]::DoEvents()
    try {
      & $so --headless --convert-to docx --outdir $f.DirectoryName $f.FullName | Out-Null
      if (Test-Path -LiteralPath ([IO.Path]::ChangeExtension($f.FullName, '.docx'))) { $ok++ } else { $fail += $f.FullName }
    } catch { $fail += $f.FullName }
  }
}
$bar.Value = $files.Count; $form.Close()

# 5. Báo kết quả
$text = "Đã đổi xong $ok / $($files.Count) file (dùng $engine).`nBản .docx nằm ngay cạnh file .doc gốc, file gốc giữ nguyên.`n`nBây giờ kéo cả thư mục vào trang Trợ lý Báo Giảng là được."
if ($fail.Count) { $text += "`n`nKhông đổi được $($fail.Count) file (có thể bị đặt mật khẩu hoặc hỏng):`n" + (($fail | Select-Object -First 8) -join "`n") }
Msg $text ($(if ($fail.Count) { 'Warning' } else { 'Information' }))
