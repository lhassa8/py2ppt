"""Tests for video and audio functionality."""

import tempfile
from pathlib import Path

import pytest

import py2ppt as ppt


class TestAddVideo:
    """Tests for add_video function."""

    def test_add_video_file_not_found(self):
        """Test adding non-existent video file raises error."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        with pytest.raises(FileNotFoundError):
            ppt.add_video(pres, 1, "nonexistent.mp4")

    def test_add_video_unsupported_format(self):
        """Test adding unsupported video format raises error."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        # Create a temp file with unsupported extension
        with tempfile.NamedTemporaryFile(suffix=".xyz", delete=False) as f:
            f.write(b"fake video data")
            temp_path = f.name

        try:
            with pytest.raises(ValueError, match="Unsupported video format"):
                ppt.add_video(pres, 1, temp_path)
        finally:
            Path(temp_path).unlink()

    def test_add_video_with_position(self):
        """Test adding video with position returns shape ID."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        # Create a minimal MP4-like file (just testing file handling)
        with tempfile.NamedTemporaryFile(suffix=".mp4", delete=False) as f:
            f.write(b"\x00\x00\x00\x1c\x66\x74\x79\x70")  # MP4 file signature
            temp_path = f.name

        try:
            shape_id = ppt.add_video(
                pres, 1, temp_path,
                left="1in", top="1in", width="6in", height="4in"
            )
            assert shape_id > 0
        finally:
            Path(temp_path).unlink()


class TestAddAudio:
    """Tests for add_audio function."""

    def test_add_audio_file_not_found(self):
        """Test adding non-existent audio file raises error."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        with pytest.raises(FileNotFoundError):
            ppt.add_audio(pres, 1, "nonexistent.mp3")

    def test_add_audio_unsupported_format(self):
        """Test adding unsupported audio format raises error."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        # Create a temp file with unsupported extension
        with tempfile.NamedTemporaryFile(suffix=".xyz", delete=False) as f:
            f.write(b"fake audio data")
            temp_path = f.name

        try:
            with pytest.raises(ValueError, match="Unsupported audio format"):
                ppt.add_audio(pres, 1, temp_path)
        finally:
            Path(temp_path).unlink()

    def test_add_audio_with_position(self):
        """Test adding audio with position returns shape ID."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        # Create a minimal MP3-like file
        with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as f:
            f.write(b"\xff\xfb\x90\x00")  # MP3 frame header
            temp_path = f.name

        try:
            shape_id = ppt.add_audio(
                pres, 1, temp_path,
                left="0.5in", top="0.5in", width="1in", height="1in"
            )
            assert shape_id > 0
        finally:
            Path(temp_path).unlink()


class TestGetMediaShapes:
    """Tests for get_media_shapes function."""

    def test_get_media_shapes_empty(self):
        """Test getting media shapes from slide without media."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        media = ppt.get_media_shapes(pres, 1)

        assert isinstance(media, list)
        assert len(media) == 0

    def test_get_media_shapes_returns_list(self):
        """Test get_media_shapes returns correct structure."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        # Add a video
        with tempfile.NamedTemporaryFile(suffix=".mp4", delete=False) as f:
            f.write(b"\x00\x00\x00\x1c\x66\x74\x79\x70")
            video_path = f.name

        try:
            ppt.add_video(
                pres, 1, video_path,
                left="1in", top="1in", width="6in", height="4in"
            )

            media = ppt.get_media_shapes(pres, 1)

            assert isinstance(media, list)
            # Note: The actual video parsing may not work with our minimal test file
            # but the function should still return a list
        finally:
            Path(video_path).unlink()


class TestVideoAudioContentTypes:
    """Tests for video/audio content type mapping."""

    def test_video_content_types_defined(self):
        """Test that video content types are properly defined."""
        from py2ppt.oxml.ns import CONTENT_TYPE

        assert CONTENT_TYPE.MP4 == "video/mp4"
        assert CONTENT_TYPE.MOV == "video/quicktime"
        assert CONTENT_TYPE.WMV == "video/x-ms-wmv"

    def test_audio_content_types_defined(self):
        """Test that audio content types are properly defined."""
        from py2ppt.oxml.ns import CONTENT_TYPE

        assert CONTENT_TYPE.MP3 == "audio/mpeg"
        assert CONTENT_TYPE.WAV == "audio/wav"
        assert CONTENT_TYPE.M4A == "audio/mp4"


class TestVideoAudioRelTypes:
    """Tests for video/audio relationship types."""

    def test_video_rel_type_defined(self):
        """Test that video relationship type is defined."""
        from py2ppt.oxml.ns import REL_TYPE

        assert "video" in REL_TYPE.VIDEO.lower()

    def test_audio_rel_type_defined(self):
        """Test that audio relationship type is defined."""
        from py2ppt.oxml.ns import REL_TYPE

        assert "audio" in REL_TYPE.AUDIO.lower()

    def test_media_rel_type_defined(self):
        """Test that media relationship type is defined."""
        from py2ppt.oxml.ns import REL_TYPE

        assert "media" in REL_TYPE.MEDIA.lower()
