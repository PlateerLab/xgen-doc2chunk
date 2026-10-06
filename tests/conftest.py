import sys
import warnings
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))
warnings.filterwarnings("ignore", category=UserWarning)


@pytest.fixture
def processor(tmp_path):
    from xgen_doc2chunk.core.document_processor import DocumentProcessor

    return DocumentProcessor(image_directory=str(tmp_path / "images"))


@pytest.fixture
def make_processor(tmp_path):
    from xgen_doc2chunk.core.document_processor import DocumentProcessor

    def _make(config=None):
        return DocumentProcessor(config=config, image_directory=str(tmp_path / "images"))

    return _make


def extract(processor, path, **chunk_kwargs):
    """Extract and chunk the way xgen-documents does on upload."""
    ext = Path(path).suffix.lstrip(".")
    text = processor.extract_text(str(path), file_extension="." + ext, extract_metadata=True)
    kwargs = dict(chunk_size=1000, chunk_overlap=200, file_extension=ext,
                  preserve_tables=True, include_position_metadata=True)
    kwargs.update(chunk_kwargs)
    return text, processor.chunk_text(text, **kwargs)
