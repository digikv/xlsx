#!/bin/sh
# Builds the XLSXClass samples: xlsxwriter, xlsxreader, xlsxreport, xlsxchart
#   ./build.sh              with the default C compiler of hbmk2
#   ./build.sh -comp=clang  other hbmk2 options are passed on
# HBMK2  hbmk2 to use (default: hbmk2 from the PATH)
cd "$(dirname "$0")" || exit 1
HBMK2=${HBMK2:-hbmk2}
FAILED=
for p in xlsxwriter xlsxreader xlsxreport xlsxchart; do
    echo
    echo "=== $p"
    "$HBMK2" "$p.hbp" "$@" || FAILED="$FAILED $p"
done
echo
if [ -n "$FAILED" ]; then
    echo "Not built:$FAILED"
    exit 1
fi
echo "Built: xlsxwriter xlsxreader xlsxreport xlsxchart"
