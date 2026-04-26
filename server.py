"""
DocMatcher – Tax Document Reconciliation
========================================

Flask web layer. All business logic lives in the `core` package; this file
is intentionally thin and only handles HTTP wiring, serialisation, and
streaming.

Run:
    python server.py

Then open http://localhost:7860 in your browser.
"""

from __future__ import annotations

import json
import logging
import traceback
from io import BytesIO

import pytesseract
from flask import (
    Flask,
    Response,
    abort,
    jsonify,
    request,
    send_file,
    send_from_directory,
)

from core.config import DEBUG_PORT, OCR_LANGUAGES, PORT, SESSION_TTL_HOURS, UPLOAD_DIR
from core.csv_parser import CsvParser
from core.export import ExportBuilder
from core.matching import MatchingService
from core.ocr import OcrService
from core.sessions import SessionStore


# ── Remote debugger ──────────────────────────────────────────────────────

if DEBUG_PORT:
    import debugpy
    debugpy.listen(("0.0.0.0", DEBUG_PORT))
    log_early = logging.getLogger("docmatcher")
    print(f"⏳ debugpy listening on port {DEBUG_PORT} — attach Cursor now, or press F5")
    # Uncomment the line below to pause startup until debugger is attached:
    # debugpy.wait_for_client()


# ── Logging ──────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-5s  %(name)s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("docmatcher")


# ── App + service wiring ─────────────────────────────────────────────────

app = Flask(__name__, static_folder="static")

csv_parser = CsvParser()
ocr_service = OcrService()
matching_service = MatchingService()
session_store = SessionStore(UPLOAD_DIR, SESSION_TTL_HOURS)
export_builder = ExportBuilder(session_store)


# ── Helpers ──────────────────────────────────────────────────────────────

def _ndjson(obj: dict) -> str:
    return json.dumps(obj) + "\n"


# ── Routes ───────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return send_from_directory("static", "index.html")


@app.post("/api/match")
def api_match():
    """
    Streams progress events as newline-delimited JSON (NDJSON) while OCR runs,
    so the frontend can show a live progress bar. The final line is either
    {"type":"result", ...} on success or {"type":"error", ...} on failure.
    """
    csv_file = request.files.get("csvFile")
    receipt_files = request.files.getlist("receiptFiles")

    if not csv_file:
        return jsonify({"error": "csvFile is required"}), 400
    if not receipt_files:
        return jsonify({"error": "At least one receipt file is required"}), 400

    # Read uploads fully inside the request context — the streaming generator
    # below runs after the request has been consumed, so `request.files` would
    # no longer be accessible there.
    csv_filename = csv_file.filename
    csv_bytes = csv_file.read()
    uploaded: list[tuple[str, bytes]] = [
        (f.filename, f.read()) for f in receipt_files
    ]

    def stream():
        try:
            yield _ndjson({"type": "progress", "stage": "parsing_csv"})
            log.info("Parsing CSV: %s", csv_filename)
            transactions = csv_parser.parse(csv_bytes)

            if not transactions:
                yield _ndjson({
                    "type": "error",
                    "error": "No valid transactions found in CSV. "
                             "Check that column headers include a Date and "
                             "Amount (or Debit/Credit).",
                })
                return

            session_id, session_dir = session_store.create_session()
            total = len(uploaded)
            log.info("Parsed %d transactions. Running OCR on %d receipts (session %s)…",
                     len(transactions), total, session_id[:8])

            yield _ndjson({
                "type": "progress", "stage": "ocr_start",
                "transactions": len(transactions), "total": total,
            })

            receipts = []
            for idx, (filename, content) in enumerate(uploaded, start=1):
                stored_name = session_store.save_file(session_dir, filename, content)
                receipts.append(ocr_service.extract(content, stored_name))
                yield _ndjson({
                    "type": "progress", "stage": "ocr",
                    "current": idx, "total": total, "file": filename,
                })

            yield _ndjson({"type": "progress", "stage": "matching"})
            log.info("OCR complete. Running matching…")
            results = matching_service.match(transactions, receipts)

            matched_count = sum(1 for r in results if r["confidence"] != "None")
            log.info("Done. Matched %d of %d transactions.", matched_count, len(transactions))

            removed = session_store.cleanup_expired()
            if removed:
                log.info("Cleaned up %d expired sessions.", removed)

            yield _ndjson({
                "type": "result",
                "sessionId": session_id,
                "results": results,
            })

        except Exception as exc:
            log.error("Request failed: %s", exc)
            traceback.print_exc()
            yield _ndjson({"type": "error", "error": str(exc)})

    # X-Accel-Buffering disables proxy buffering (nginx) so events stream live.
    return Response(
        stream(),
        mimetype="application/x-ndjson",
        headers={"X-Accel-Buffering": "no", "Cache-Control": "no-cache"},
    )


@app.get("/api/receipt/<session_id>/<path:filename>")
def api_receipt(session_id: str, filename: str):
    """Serves a receipt file from a specific session for in-browser preview."""
    filepath = session_store.get_file_path(session_id, filename)
    if filepath is None:
        abort(404)
    # Inline disposition so PDFs/images render in the browser instead of downloading
    return send_from_directory(
        str(filepath.parent),
        filepath.name,
        as_attachment=False,
    )


@app.post("/api/export")
def api_export():
    """
    Generates a ZIP archive with renamed receipts + overview.xlsx.
    The frontend sends the current results including any manual assignments.
    """
    try:
        payload = request.get_json(silent=True) or {}
        session_id = payload.get("sessionId")
        results = payload.get("results", [])

        if not session_id or not results:
            return jsonify({"error": "sessionId and results are required"}), 400

        log.info("Building export for session %s (%d transactions)…",
                 session_id[:8], len(results))

        zip_bytes, filename = export_builder.build(session_id, results)

        log.info("Export ready: %s (%.1f KB)", filename, len(zip_bytes) / 1024)
        return send_file(
            BytesIO(zip_bytes),
            mimetype="application/zip",
            as_attachment=True,
            download_name=filename,
        )

    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        log.error("Export failed: %s", exc)
        traceback.print_exc()
        return jsonify({"error": str(exc)}), 500


# ── Entry point ──────────────────────────────────────────────────────────

def _check_tesseract() -> None:
    try:
        version = pytesseract.get_tesseract_version()
        log.info("Tesseract %s ready. Languages: %s", version, OCR_LANGUAGES)
    except Exception as exc:
        log.warning("Tesseract not found: %s", exc)
        log.warning("Install it first:")
        log.warning("  macOS:  brew install tesseract")
        log.warning("  Ubuntu: sudo apt install tesseract-ocr")
        log.warning("  Windows: https://github.com/UB-Mannheim/tesseract/wiki")


if __name__ == "__main__":
    _check_tesseract()
    log.info("Starting DocMatcher on http://localhost:%d", PORT)
    app.run(host="0.0.0.0", port=PORT, debug=False)
