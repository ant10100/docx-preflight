"""Read-only DOCX inspection; findings are hints, not a safety certificate."""

from .scanner import inspect_docx

__all__ = ['inspect_docx']
__version__ = '0.1.0'
