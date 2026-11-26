#!/bin/sh
# usage: run-office2pdf.sh xlsx|pptx PAIRS.jsonl OUTDIR   (runs on fire; copies inputs to the Windows PC, runs desktop Excel/PowerPoint in the
# interactive session via scheduled task eo-excel2pdf / eo-ppt2pdf, fetches <id>.pdf + <id>.json into OUTDIR, removes the PC input copies)
set -u
K=$HOME/.ssh/id_ed25519_windows_qa; H=${WINDOWS_QA_HOST}
O="-i $K -o UserKnownHostsFile=/tmp/kh_pc -o StrictHostKeyChecking=accept-new -o BatchMode=yes"
W() { ssh $O $H "$@" 2>&1 | grep -v -e '^\*\*'; }
kind=$1; pairs=$2; out=$3; mkdir -p "$out"
if [ "$kind" = xlsx ]; then I=inx; Ob=outx; T=eo-excel2pdf; else I=inp; Ob=outp; T=eo-ppt2pdf; fi
W "cmd /c mkdir C:\\eo-ref\\$I C:\\eo-ref\\$Ob 2>nul"
W "cmd /c del /q C:\\eo-ref\\$I\\*.* C:\\eo-ref\\$Ob\\*.* 2>nul"
python3 -I -c "import json,sys
for l in open(sys.argv[1]):
    d=json.loads(l); print(d['input'],d['id'])" "$pairs" | while read f n; do
  scp $O "$f" "$H:C:/eo-ref/$I/$n.$kind" 2>&1 | grep -v -e '^\*\*'; done
W "schtasks /Run /TN $T"
i=0; until W "cmd /c if exist C:\\eo-ref\\$Ob\\done.marker echo DONE" | grep -q DONE; do i=$((i+1)); [ $i -gt 360 ] && { echo TIMEOUT; break; }; sleep 5; done
scp $O "$H:C:/eo-ref/$Ob/*" "$out/" 2>&1 | grep -v -e '^\*\*'
W "cmd /c del /q C:\\eo-ref\\$I\\*.* 2>nul"
tail -3 "$out/run.log"
