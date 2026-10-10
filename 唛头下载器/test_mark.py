import base64
import json
import unittest
from unittest.mock import patch

import requests
import app


PDF = b"%PDF-1.4\nfixture\n%%EOF"


def response(content, content_type="application/json"):
    result = requests.Response()
    result.status_code = 200
    result._content = content
    result.headers["Content-Type"] = content_type
    return result


class MarkTests(unittest.TestCase):
    def test_download_uses_new_mark_endpoint(self):
        def get(url, **kwargs):
            if url != "http://newapi.ywcd56.com/openapi/getOutMarkPdf":
                raise requests.ConnectionError("old mark endpoint")
            self.assertEqual(kwargs["params"], {"waybill_num": "TEST0001"})
            self.assertEqual(kwargs["headers"], {"x-api-key": "test-key", "token": "test-token"})
            return response(PDF, "application/pdf")

        with patch.object(app.requests, "get", side_effect=get):
            self.assertEqual(app.fetch_mark("TEST0001", "test-key", "test-token"), (PDF, None))

    def test_relative_pdf_path_uses_new_mark_host(self):
        def get(url, **kwargs):
            if url.endswith("/openapi/getOutMarkPdf"):
                return response(b'{"code":200,"path":"/files/mark.pdf"}')
            if url == "http://newapi.ywcd56.com/files/mark.pdf":
                return response(PDF, "application/pdf")
            raise requests.ConnectionError("wrong PDF host")

        with patch.object(app.requests, "get", side_effect=get):
            self.assertEqual(app.fetch_mark("TEST0001", "test-key", "test-token"), (PDF, None))

    def test_base64_path_is_decoded_without_second_request(self):
        payload = json.dumps({"code": 200, "path": base64.b64encode(PDF).decode()}).encode()
        with patch.object(app.requests, "get", side_effect=[response(payload), requests.ConnectionError("base64 is not a URL")]):
            self.assertEqual(app.fetch_mark("TEST0001", "test-key", "test-token"), (PDF, None))

    def test_pdf_data_uri_is_decoded(self):
        payload = json.dumps({"code": 200, "path": "data:application/pdf;base64," + base64.b64encode(PDF).decode()}).encode()
        with patch.object(app.requests, "get", side_effect=[response(payload)]):
            self.assertEqual(app.fetch_mark("test-waybill", "test-key", "test-token"), (PDF, None))

    def test_absolute_pdf_url_still_downloads(self):
        def get(url, **kwargs):
            if url.endswith("/openapi/getOutMarkPdf"):
                return response(b'{"code":200,"path":"http://newapi.ywcd56.com/custom.pdf"}')
            if url == "http://newapi.ywcd56.com/custom.pdf":
                return response(PDF, "application/pdf")
            raise requests.ConnectionError("unexpected URL")
        with patch.object(app.requests, "get", side_effect=get):
            self.assertEqual(app.fetch_mark("test-waybill", "test-key", "test-token"), (PDF, None))

    def test_non_pdf_data_uri_is_rejected(self):
        payload = json.dumps({"code": 200, "path": "data:application/pdf;base64,aGVsbG8="}).encode()
        with patch.object(app.requests, "get", side_effect=[response(payload)]):
            pdf, error = app.fetch_mark("test-waybill", "test-key", "test-token")
        self.assertIsNone(pdf)
        self.assertEqual(error, "接口返回的Base64内容不是PDF")

    def test_business_error_is_preserved(self):
        payload = json.dumps({"code": 400, "msg": "暂无数据！"}).encode()
        with patch.object(app.requests, "get", return_value=response(payload)):
            pdf, error = app.fetch_mark("test-waybill", "test-key", "test-token")
        self.assertIsNone(pdf)
        self.assertIn("暂无数据", error)


if __name__ == "__main__":
    unittest.main()
