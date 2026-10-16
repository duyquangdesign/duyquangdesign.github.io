#!/bin/bash
# Trợ lý Báo Giảng – Đổi file Word .doc sang .docx (macOS)
# Cách chạy: chuột phải vào file này > Mở (Open). Nếu không mở được: mở Terminal, gõ "bash " rồi kéo file này vào, nhấn Enter.
TITLE="Trợ lý Báo Giảng"
say_dialog() { osascript -e "display dialog \"$1\" with title \"$TITLE\" buttons {\"OK\"} default button 1" >/dev/null 2>&1; }

ROOT=$(osascript -e 'POSIX path of (choose folder with prompt "Chọn thư mục chứa giáo án .doc (đổi cả thư mục con):")' 2>/dev/null)
[ -z "$ROOT" ] && exit 0

FILES=()
while IFS= read -r -d '' f; do
  b=$(basename "$f"); [[ "$b" == "~\$"* ]] && continue
  [ -e "${f%.*}.docx" ] && continue
  FILES+=("$f")
done < <(find "$ROOT" -type f \( -iname '*.doc' \) -print0)
TOTAL=${#FILES[@]}
if [ "$TOTAL" -eq 0 ]; then say_dialog "Không có file .doc nào cần đổi trong thư mục đã chọn."; exit 0; fi

SOFFICE="/Applications/LibreOffice.app/Contents/MacOS/soffice"
HAS_WORD=0; [ -d "/Applications/Microsoft Word.app" ] && HAS_WORD=1
OK=0; FAIL=0; USED_TEXTUTIL=0; i=0
for f in "${FILES[@]}"; do
  i=$((i+1)); out="${f%.*}.docx"
  echo "[$i/$TOTAL] $(basename "$f")"
  if [ $HAS_WORD -eq 1 ]; then
    osascript - "$f" "$out" >/dev/null 2>&1 <<'APPLESCRIPT'
on run argv
  set inF to POSIX file (item 1 of argv)
  set outP to item 2 of argv
  tell application "Microsoft Word"
    open inF
    set d to active document
    save as d file name outP file format format document
    close d saving no
  end tell
end run
APPLESCRIPT
  fi
  if [ ! -e "$out" ] && [ -x "$SOFFICE" ]; then
    "$SOFFICE" --headless --convert-to docx --outdir "$(dirname "$f")" "$f" >/dev/null 2>&1
  fi
  if [ ! -e "$out" ]; then
    textutil -convert docx -output "$out" "$f" >/dev/null 2>&1 && USED_TEXTUTIL=1
  fi
  if [ -e "$out" ]; then OK=$((OK+1)); else FAIL=$((FAIL+1)); fi
done
[ $HAS_WORD -eq 1 ] && osascript -e 'tell application "Microsoft Word" to quit saving no' >/dev/null 2>&1
MSG="Đã đổi xong $OK / $TOTAL file. Bản .docx nằm cạnh file .doc gốc, file gốc giữ nguyên."
[ $FAIL -gt 0 ] && MSG="$MSG Có $FAIL file không đổi được."
[ $USED_TEXTUTIL -eq 1 ] && MSG="$MSG Lưu ý: máy chưa có Word/LibreOffice nên một số file có thể mất hình ảnh – nên cài LibreOffice (miễn phí) rồi chạy lại."
say_dialog "$MSG"
