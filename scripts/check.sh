#!/bin/sh
# Static checks for the exported VBA source. This is the VBA equivalent of a unit-test run:
# it cannot compile or execute VBA (that needs Access/Excel/Word), it only catches the
# mistakes that can be found by reading the text. Run from anywhere: sh scripts/check.sh
#
# Exit status 1 if any check fails.

root=$(git rev-parse --show-toplevel 2>/dev/null) || root=$(cd "$(dirname "$0")/.." && pwd)
cd "$root" || exit 1

fail=0
err() { echo "FAIL: $1"; fail=1; }

files=$(find src test -type f \( -name '*.bas' -o -name '*.cls' -o -name '*.frm' \) | sort)
[ -z "$files" ] && { echo "check: no VBA source files found under src/"; exit 0; }

for f in $files; do
  base=$(basename "$f"); stem=${base%.*}

  # 1. Option Explicit must be declared (undeclared variables become silent Variants)
  grep -q '^Option Explicit' "$f" || err "$f: missing 'Option Explicit'"

  # 2. The VB_Name attribute must match the file name, or an import renames the component
  name=$(sed -n 's/^Attribute VB_Name = "\(.*\)".*/\1/p' "$f" | tr -d '\r' | head -1)
  [ -z "$name" ] && err "$f: missing 'Attribute VB_Name'"
  [ -n "$name" ] && [ "$name" != "$stem" ] && err "$f: VB_Name \"$name\" does not match file name \"$stem\""

  # 3. Merge conflict markers
  grep -nE '^(<<<<<<<|=======$|>>>>>>>)' "$f" >/dev/null && err "$f: merge conflict marker"

  # 4. API declarations must be PtrSafe (64-bit Office will not compile them otherwise)
  grep -nEi '^[[:space:]]*(Public |Private )?Declare +(Function|Sub)' "$f" | grep -vi 'PtrSafe' >/dev/null \
    && err "$f: Declare statement without PtrSafe"

  # 5. Mixed line endings (a file should be all CRLF or all LF)
  crlf=$(grep -c "$(printf '\r')$" "$f"); total=$(wc -l < "$f")
  [ "$crlf" -ne 0 ] && [ "$crlf" -ne "$total" ] && err "$f: mixed CRLF and LF line endings"

  # 6. Anything that looks like a credential or a client connection string (this repo is public)
  grep -nEi '(password|pwd|uid)[[:space:]]*=[[:space:]]*[^[:space:]"&]+|Data Source=|Server=|\\\\[A-Za-z0-9_.-]+\\|[A-Za-z]:\\Users\\' "$f" >/dev/null \
    && err "$f: possible credential, connection string or personal/client path (review by hand)"
done

# 7. No host files or lock files committed
bad=$(git ls-files 2>/dev/null | grep -Ev '^test/[^/]*\.xlsm$' | grep -Ei '\.(accdb|accde|mdb|xlsm|xlsb|xlam|docm|dotm|laccdb|ldb)$|(^|/)~\$')
[ -n "$bad" ] && err "host or lock file committed: $bad"

if [ "$fail" -eq 0 ]; then echo "check: all static checks passed ($(echo "$files" | wc -l | tr -d ' ') files)."; fi
exit $fail
