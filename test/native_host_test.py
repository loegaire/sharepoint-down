import importlib.util
from pathlib import Path
import unittest


MODULE_PATH = Path(__file__).resolve().parents[1] / "native" / "sp_downloader_host.py"
spec = importlib.util.spec_from_file_location("sp_downloader_host", MODULE_PATH)
host = importlib.util.module_from_spec(spec)
spec.loader.exec_module(host)


class NativeHostTest(unittest.TestCase):
    def test_sanitizes_download_filename(self):
        self.assertEqual(host.sanitize_filename("../Bad: File.pdf"), "Bad_File.pdf")
        self.assertEqual(host.sanitize_filename(""), "sharepoint-document.pdf")

    def test_validates_sharepoint_download_request(self):
        message = {
            "url": "https://husteduvn-my.sharepoint.com/personal/dat/_layouts/15/download.aspx?SourceUrl=%2Fpersonal%2Fdat%2FDocuments%2Fa.pdf",
            "filename": "a.pdf",
            "cookieHeader": "FedAuth=fed; rtFa=rt",
        }

        host.validate_message(message)

    def test_rejects_missing_cookies_and_non_sharepoint_hosts(self):
        with self.assertRaises(ValueError):
            host.validate_message({
                "url": "https://example.com/download.aspx",
                "filename": "a.pdf",
                "cookieHeader": "FedAuth=fed; rtFa=rt",
            })

        with self.assertRaises(ValueError):
            host.validate_message({
                "url": "https://husteduvn-my.sharepoint.com/_layouts/15/download.aspx?SourceUrl=%2Fa.pdf",
                "filename": "a.pdf",
                "cookieHeader": "FedAuth=fed",
            })


if __name__ == "__main__":
    unittest.main()
