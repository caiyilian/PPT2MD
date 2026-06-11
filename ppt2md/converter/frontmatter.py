"""YAML frontmatter generation from document properties."""


def extract_document_properties(prs):
    """Extract core properties from a Presentation object.

    Args:
        prs: python-pptx Presentation object.

    Returns:
        dict with title, author, created, modified, source.
    """
    props = prs.core_properties
    return {
        "title": props.title or "",
        "author": props.author or "",
        "created": str(props.created) if props.created else "",
        "modified": str(props.modified) if props.modified else "",
        "subject": props.subject or "",
        "keywords": props.keywords or "",
        "category": props.category or "",
        "comments": props.comments or "",
    }


def generate_frontmatter(prs, source_file=""):
    """Generate YAML frontmatter string from presentation properties.

    Args:
        prs: python-pptx Presentation object.
        source_file: Original filename.

    Returns:
        str: YAML frontmatter block.
    """
    props = extract_document_properties(prs)
    lines = ["---"]

    if props["title"]:
        lines.append('title: "{}"'.format(props["title"].replace('"', '\\"')))
    if props["author"]:
        lines.append('author: "{}"'.format(props["author"].replace('"', '\\"')))
    if source_file:
        lines.append('source: "{}"'.format(source_file))
    if props["created"]:
        lines.append('created: "{}"'.format(props["created"]))
    if props["modified"]:
        lines.append('modified: "{}"'.format(props["modified"]))
    if props["subject"]:
        lines.append('subject: "{}"'.format(props["subject"].replace('"', '\\"')))
    if props["keywords"]:
        lines.append('keywords: "{}"'.format(props["keywords"].replace('"', '\\"')))

    lines.append("---")
    return "\n".join(lines)
