"""XML namespaces for PresentationML (Open XML)."""


# Core PresentationML namespaces
NAMESPACES: dict[str, str] = {
    # PresentationML
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "p15": "http://schemas.microsoft.com/office/powerpoint/2012/main",

    # DrawingML
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",

    # Relationships
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",

    # Package relationships
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",

    # Content types
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",

    # Core properties
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",

    # Extended properties
    "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",

    # XML base
    "xml": "http://www.w3.org/XML/1998/namespace",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",

    # VML (legacy shapes)
    "v": "urn:schemas-microsoft-com:vml",

    # Office
    "o": "urn:schemas-microsoft-com:office:office",

    # Chart
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",

    # Diagram
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",

    # Picture
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
}

# For convenience, also create reverse mapping
PREFIXES: dict[str, str] = {v: k for k, v in NAMESPACES.items()}

# Default namespace map for lxml
nsmap = NAMESPACES.copy()


def qn(tag: str) -> str:
    """Convert a prefixed tag name to Clark notation.

    Example:
        qn("p:sld") -> "{http://schemas.openxmlformats.org/presentationml/2006/main}sld"
        qn("a:t") -> "{http://schemas.openxmlformats.org/drawingml/2006/main}t"

    Args:
        tag: Tag name with namespace prefix (e.g., "p:sld")

    Returns:
        Tag in Clark notation (e.g., "{namespace}localname")
    """
    if ":" not in tag:
        return tag

    prefix, local = tag.split(":", 1)
    if prefix not in NAMESPACES:
        raise ValueError(f"Unknown namespace prefix: {prefix}")

    return f"{{{NAMESPACES[prefix]}}}{local}"


def local_name(tag: str) -> str:
    """Extract local name from a Clark notation tag.

    Example:
        local_name("{http://...}sld") -> "sld"

    Args:
        tag: Tag in Clark notation

    Returns:
        Local name without namespace
    """
    if tag.startswith("{"):
        return tag.split("}", 1)[1]
    return tag


def prefix_for_namespace(namespace: str) -> str | None:
    """Get the prefix for a namespace URI.

    Args:
        namespace: Namespace URI

    Returns:
        Prefix string or None if not found
    """
    return PREFIXES.get(namespace)


# Relationship types
class REL_TYPE:
    """Relationship type URIs used in .rels files."""

    OFFICE_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
    SLIDE_LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
    SLIDE_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
    THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
    NOTES_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
    NOTES_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster"
    HANDOUT_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster"
    PRES_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps"
    VIEW_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps"
    TABLE_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles"
    CORE_PROPS = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
    EXT_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
    IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
    TAGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags"


# Content types
class CONTENT_TYPE:
    """Content type strings for [Content_Types].xml."""

    PRESENTATION = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
    PRESENTATION_MACRO = "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"
    TEMPLATE = "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
    SLIDESHOW = "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml"

    SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
    SLIDE_LAYOUT = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
    SLIDE_MASTER = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"

    NOTES_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"
    NOTES_MASTER = "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"
    HANDOUT_MASTER = "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml"

    THEME = "application/vnd.openxmlformats-officedocument.theme+xml"

    PRES_PROPS = "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"
    VIEW_PROPS = "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"
    TABLE_STYLES = "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"

    CORE_PROPS = "application/vnd.openxmlformats-package.core-properties+xml"
    EXT_PROPS = "application/vnd.openxmlformats-officedocument.extended-properties+xml"

    RELS = "application/vnd.openxmlformats-package.relationships+xml"
    XML = "application/xml"

    # Images
    PNG = "image/png"
    JPEG = "image/jpeg"
    GIF = "image/gif"
    BMP = "image/bmp"
    TIFF = "image/tiff"
    SVG = "image/svg+xml"
    EMF = "image/x-emf"
    WMF = "image/x-wmf"
