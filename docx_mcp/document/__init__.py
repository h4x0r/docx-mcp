"""DocxDocument: mixin composition and public API."""

from .accessibility import AccessibilityMixin
from .base import (
    CP,
    CT,
    CT_TYPES,
    DC,
    DCTERMS,
    NSMAP,
    REL_TYPES,
    RELS,
    W14,
    W15,
    WP,
    XML_SPACE,
    A,
    BaseMixin,
    R,
    V,
    W,
    _now_iso,
    _preserve,
)
from .bookmarks import BookmarksMixin
from .captions import CaptionMixin
from .charts import ChartsMixin
from .clausediff import ClauseDiffMixin
from .comments import CommentsMixin
from .compare import CompareMixin
from .contentcontrols import ContentControlsMixin
from .creation import CreationMixin
from .endnotes import EndnotesMixin
from .equations import EquationsMixin
from .fields import FieldsMixin
from .footnotes import FootnotesMixin
from .formatting import FormattingMixin
from .headers_footers import HeadersFootersMixin
from .hyperlinks import HyperlinksMixin
from .images import ImagesMixin
from .lists import ListsMixin
from .litigation import LitigationMixin
from .markdown_export import MarkdownExportMixin
from .merge import MergeMixin
from .metadata import MetadataMixin
from .pdfexport import PdfExportMixin
from .pii import PiiMixin
from .properties import PropertiesMixin
from .protection import ProtectionMixin
from .query import XPathMixin
from .rawparts import RawPartsMixin
from .reading import ReadingMixin
from .references import ReferencesMixin
from .reviewmerge import ReviewMergeMixin
from .revisions import RevisionsMixin
from .sections import SectionsMixin
from .sessionlog import SessionLogMixin
from .splitting import SplittingMixin
from .statistics import StatisticsMixin
from .styles import StylesMixin
from .tables import TablesMixin
from .template import TemplateMixin
from .textboxes import TextBoxesMixin
from .theme import ThemeMixin
from .toc import TocMixin
from .tracks import TracksMixin
from .validation import ValidationMixin
from .watermark import WatermarkMixin


class DocxDocument(
    BaseMixin,
    CreationMixin,
    BookmarksMixin,
    ContentControlsMixin,
    ReadingMixin,
    TracksMixin,
    RevisionsMixin,
    FormattingMixin,
    CommentsMixin,
    FootnotesMixin,
    ValidationMixin,
    TablesMixin,
    StylesMixin,
    HeadersFootersMixin,
    ListsMixin,
    PropertiesMixin,
    FieldsMixin,
    HyperlinksMixin,
    ImagesMixin,
    EndnotesMixin,
    SectionsMixin,
    ReferencesMixin,
    ProtectionMixin,
    MergeMixin,
    MetadataMixin,
    CompareMixin,
    PiiMixin,
    RawPartsMixin,
    XPathMixin,
    TocMixin,
    TemplateMixin,
    LitigationMixin,
    EquationsMixin,
    ChartsMixin,
    ReviewMergeMixin,
    ClauseDiffMixin,
    MarkdownExportMixin,
    SessionLogMixin,
    ThemeMixin,
    CaptionMixin,
    WatermarkMixin,
    SplittingMixin,
    StatisticsMixin,
    AccessibilityMixin,
    TextBoxesMixin,
    PdfExportMixin,
):
    """Word document editor with OOXML-level control."""

    pass


__all__ = [
    "DocxDocument",
    "W",
    "W14",
    "W15",
    "R",
    "V",
    "A",
    "CT",
    "RELS",
    "XML_SPACE",
    "WP",
    "DC",
    "DCTERMS",
    "CP",
    "NSMAP",
    "REL_TYPES",
    "CT_TYPES",
    "_now_iso",
    "_preserve",
    "WatermarkMixin",
    "SplittingMixin",
]
