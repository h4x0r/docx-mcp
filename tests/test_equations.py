import pytest
from docx_mcp.document import DocxDocument, W, W14
from docx_mcp.document.errors import DocxMcpError, ErrCode

M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"


def _make_doc(tmp_path):
    return DocxDocument.create(str(tmp_path / "test.docx"))


def _get_para_id(doc):
    tree = doc._tree("word/document.xml")
    paras = tree.findall(f".//{W}p")
    return paras[0].get(f"{W14}paraId")


class TestEquations:
    def test_add_simple_equation(self, tmp_path):
        """add_equation inserts oMathPara after paragraph."""
        pytest.importorskip("latex2mathml")
        doc = _make_doc(tmp_path)
        para_id = _get_para_id(doc)
        result = doc.add_equation(para_id, r"x^2")
        assert result["para_id"] == para_id
        assert result["latex"] == r"x^2"

    def test_omml_in_document_xml(self, tmp_path):
        """After add_equation, m:oMath exists in document XML."""
        pytest.importorskip("latex2mathml")
        doc = _make_doc(tmp_path)
        para_id = _get_para_id(doc)
        doc.add_equation(para_id, r"\frac{1}{2}")
        tree = doc._tree("word/document.xml")
        omaths = list(tree.iter(f"{M}oMath"))
        assert len(omaths) >= 1

    def test_get_equations_roundtrip(self, tmp_path):
        """get_equations returns the equation after add_equation."""
        pytest.importorskip("latex2mathml")
        doc = _make_doc(tmp_path)
        para_id = _get_para_id(doc)
        doc.add_equation(para_id, r"E = mc^2")
        equations = doc.get_equations()
        assert len(equations) >= 1
        assert "omml_xml" in equations[0]
        assert f"{M}oMath" in equations[0]["omml_xml"] or "oMath" in equations[0]["omml_xml"]

    def test_missing_dep_graceful_error(self, tmp_path, monkeypatch):
        """add_equation raises DocxMcpError when latex2mathml is unavailable."""
        # Simulate missing module by patching the import
        monkeypatch.setattr("builtins.__import__", _make_import_blocker("latex2mathml"))
        doc = _make_doc(tmp_path)
        para_id = _get_para_id(doc)
        with pytest.raises(DocxMcpError) as exc_info:
            doc.add_equation(para_id, r"x^2")
        assert exc_info.value.code == ErrCode.PII_DEPS_MISSING


def _make_import_blocker(blocked_module):
    real_import = __import__
    def patched_import(name, *args, **kwargs):
        if name == blocked_module or name.startswith(blocked_module + "."):
            raise ImportError(f"Mocked: {name} not available")
        return real_import(name, *args, **kwargs)
    return patched_import
