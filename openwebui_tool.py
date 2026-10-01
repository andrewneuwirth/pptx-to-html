"""
title: PowerPoint Viewer
description: Show an uploaded .pptx inside the chat as a slide deck with Previous / Next controls.
requirements: python-pptx
version: 0.1.0
"""
# Open WebUI Tool: returns the deck as a Rich UI embed (HTMLResponse + Content-Disposition: inline).
# pptx_to_html.py must be importable by the Open WebUI backend (same site-packages / PYTHONPATH).
# NOTE: the file lookup below (Files / Storage) follows Open WebUI's internal API and has not been
# run against a live instance; adjust it to your version if uploads are stored differently.
import json

from fastapi.responses import HTMLResponse

from pptx_to_html import convert, slide_text


def _pptx_path(files):
    from open_webui.models.files import Files
    from open_webui.storage.provider import Storage

    for f in files or []:
        meta = f.get("file") or f
        name = (meta.get("filename") or meta.get("name") or f.get("name") or "").lower()
        if not name.endswith(".pptx"):
            continue
        rec = Files.get_file_by_id(meta.get("id") or f.get("id"))
        if rec is not None and rec.path:
            return Storage.get_file(rec.path), name
    return None, None


class Tools:
    def show_presentation(self, __files__: list = None) -> HTMLResponse:
        """
        Display the PowerPoint (.pptx) file attached to this chat as an interactive slide viewer.
        Use when the user asks to see, open, preview or present an uploaded presentation.
        """
        path, name = _pptx_path(__files__)
        if path is None:
            return "No .pptx file is attached to this chat."
        html = convert(path, viewer="deck", max_image_px=1600, title=name)
        slides = slide_text(path)
        context = json.dumps({"status": "shown", "file": name, "slide_count": len(slides), "slides": slides})[:60000]
        return HTMLResponse(content=html, headers={"Content-Disposition": "inline"}), context
