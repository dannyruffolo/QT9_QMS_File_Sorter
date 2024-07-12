import pytest
import requests_mock
from main import check_for_updates  # Adjust the import path as necessary

@pytest.fixture
def mock_response():
	return {
		"tag_name": "2.1.1",
		"assets": [{"browser_download_url": "https://example.com/download"}]
	}

def test_check_for_updates_with_newer_version_available(mock_response):
	with requests_mock.Mocker() as m:
		m.get("https://api.github.com/repos/dannyruffolo/QT9_QMS_File_Sorter/releases/latest", json=mock_response)
		update_available, latest_version, download_url = check_for_updates()
		assert update_available is True
		assert latest_version == "2.1.2"
		assert download_url == "https://example.com/download"

def test_check_for_updates_with_no_update_available(mock_response):
	mock_response["tag_name"] = "1.0.0"  # Same as current version
	with requests_mock.Mocker() as m:
		m.get("https://api.github.com/repos/dannyruffolo/QT9_QMS_File_Sorter/releases/latest", json=mock_response)
		update_available, latest_version, download_url = check_for_updates()
		assert update_available is False
		assert latest_version is None
		assert download_url is None