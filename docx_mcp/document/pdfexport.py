"""PDF export mixin: convert the open document to PDF via LibreOffice headless."""
from __future__ import annotations

import shutil
import subprocess
from pathlib import Path


class PdfExportMixin:

    def convert_to_pdf(self, output_path: str) -> dict:
        """Convert the current document to PDF using LibreOffice headless.

        Saves the document first, then invokes:
          libreoffice --headless --convert-to pdf --outdir <dir> <docx>

        Args:
            output_path: Desired path for the output PDF file.

        Returns:
            {"pdf_path": str}

        Raises:
            RuntimeError: If no document is open, LibreOffice is not found,
                          or the conversion process exits with a non-zero code.
        """
        if self.workdir is None:
            raise RuntimeError("No document is open.")

        lo = shutil.which("libreoffice") or shutil.which("soffice")
        if lo is None:
            raise RuntimeError(
                "LibreOffice not found. Install it and ensure 'libreoffice' or "
                "'soffice' is on PATH."
            )

        self.save(self.source_path, backup=False)

        out = Path(output_path)
        outdir = out.parent
        outdir.mkdir(parents=True, exist_ok=True)

        result = subprocess.run(
            [lo, "--headless", "--convert-to", "pdf", "--outdir", str(outdir),
             str(self.source_path)],
            capture_output=True,
            text=True,
        )
        if result.returncode != 0:
            raise RuntimeError(
                f"LibreOffice conversion failed (exit {result.returncode}): "
                f"{result.stderr.strip()}"
            )

        # LibreOffice names the output after the input stem
        generated = outdir / (Path(self.source_path).stem + ".pdf")
        if generated != out and generated.exists():
            generated.rename(out)

        return {"pdf_path": str(out)}
