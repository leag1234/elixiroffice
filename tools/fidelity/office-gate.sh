#!/usr/bin/env bash
# SPDX-License-Identifier: AGPL-3.0-only
# One-command spreadsheet/presentation fidelity gate (app PDF export vs MS Office desktop reference PDFs).
# usage: office-gate.sh xlsx|pptx PAIRS.jsonl OUTDIR [APP_ROOT]
#   PAIRS.jsonl : lines with id, input, reference (PDF made by Excel/PowerPoint; see office-ref/run-office2pdf.sh)
#   APP_ROOT    : extracted AppImage (squashfs-root), default ${XPFID_APP_ROOT}
# Spreadsheets are exported as the ENTIRE WORKBOOK (EO_FIDELITY_SHEET_SCOPE=active to restore the active-sheet default).
set -eu
kind=$1; pairs=$2; out=$3; app=${4:-${XPFID_APP_ROOT}}
here=$(cd "$(dirname "$0")" && pwd)
W=${XPFID_WORK}
ln -sfn /opt/elixiroffice-qa/chrome-sandbox-44.9.1 "$app/chrome-sandbox"
cd "$here/../.."
export PYTHONPATH=$W/deps PATH=$W/root/usr/bin:$PATH XKB_BINDIR=$W/root/usr/bin
nice -n 10 ${XPFID_RENDER_DESKTOP} --manifest "$pairs" --out "$out" --executable "$app/elixiroffice" \
  --xvfb $W/root/usr/bin/Xvfb --playwright "$HOME/elixiroffice/pw/node_modules/playwright" --font-mode bundled-only \
  --node "$HOME/elixiroffice/node-v22.23.3-linux-x64/bin/node" || true
python3 -I "$here/office-summary.py" "$kind" "$out" "$out/SUMMARY.md"
