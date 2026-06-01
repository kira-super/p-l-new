"""Run this once from oefof_pl/ to patch pipeline.py in-place.

    python patch_pipeline.py
"""
import pathlib

path = pathlib.Path("oefof_pl/pipeline.py")
text = path.read_text(encoding="utf-8")

old = (
    "        all_agg_ccys = {\n"
    "            agg.ccy for agg in list(start_agg.values()) + list(end_agg.values())\n"
    "            if agg.ccy\n"
    "        }\n"
    "        missing_ccys = sorted(all_agg_ccys - set(fx_start) - set(fx_end))\n"
)

new = (
    "        # Include CCYs from trades as well as snapshots: positions opened\n"
    "        # AND closed within the YTD period (e.g. Capital Tankers MHY1096C1093,\n"
    "        # SED Energy CY0101162119 — both NOK, bought and sold in 2026) never\n"
    "        # appear in start_agg or end_agg, so their CCY is only visible in trades.\n"
    "        all_agg_ccys = {\n"
    "            agg.ccy for agg in list(start_agg.values()) + list(end_agg.values())\n"
    "            if agg.ccy\n"
    "        }\n"
    "        if not trades_df.empty and \"CCY\" in trades_df.columns:\n"
    "            all_agg_ccys |= {\n"
    "                str(c).strip().upper()\n"
    "                for c in trades_df[\"CCY\"].dropna().unique()\n"
    "                if str(c).strip().upper() not in (\"EUR\", \"USD\", \"\")\n"
    "            }\n"
    "        missing_ccys = sorted(all_agg_ccys - set(fx_start) - set(fx_end))\n"
)

if old not in text:
    # Try with \r\n line endings (Windows)
    old_crlf = old.replace("\n", "\r\n")
    new_crlf = new.replace("\n", "\r\n")
    if old_crlf in text:
        text = text.replace(old_crlf, new_crlf)
        path.write_text(text, encoding="utf-8")
        print("Patched (CRLF line endings).")
    else:
        print("ERROR: target block not found. Check the pipeline.py version.")
        raise SystemExit(1)
else:
    text = text.replace(old, new)
    path.write_text(text, encoding="utf-8")
    print("Patched (LF line endings).")

# Verify
if "trades_df[\"CCY\"]" in path.read_text(encoding="utf-8"):
    print("Verification OK — patch is in place.")
else:
    print("ERROR: patch did not apply correctly.")
