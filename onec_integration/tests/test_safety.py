import json
import os
import tempfile
import unittest
from pathlib import Path

from onec_integration.config import load_config
from onec_integration.errors import ErrorCode, IntegrationError
from onec_integration.service import OneCIntegrationService


class FakeClient:
    def query(self, text, params=None):
        return [{"quantity": 10, "reserved": 3}]


class SafetyTests(unittest.TestCase):
    def setUp(self):
        self.old_env = os.environ.copy()
        os.environ["ONEC_PASSWORD"] = "secret"

    def tearDown(self):
        os.environ.clear()
        os.environ.update(self.old_env)

    def test_write_is_blocked_without_explicit_flag(self):
        os.environ["DRY_RUN"] = "false"
        os.environ["ALLOW_DOCUMENT_WRITE"] = "false"
        config = load_config()
        service = OneCIntegrationService(config, FakeClient())

        with self.assertRaises(IntegrationError) as ctx:
            service.create_assembly({"external_id": "test-1"})

        self.assertEqual(ctx.exception.code, ErrorCode.DOCUMENT_WRITE_ERROR)

    def test_dry_run_allows_document_preview(self):
        os.environ["DRY_RUN"] = "true"
        config = load_config()
        service = OneCIntegrationService(config, FakeClient())

        response = service.create_transfer(
            {"external_id": "test-2", "source_warehouse_id": "A", "target_warehouse_id": "B"}
        )

        self.assertTrue(response["success"])
        self.assertTrue(response["dry_run"])

    def test_stock_query_requires_real_metadata_names(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "config.json"
            path.write_text(
                json.dumps(
                    {
                        "connection": {
                            "server": "pz-sql1.pzmc.org:1541",
                            "database": "UT_dev",
                            "username_env": "ONEC_USERNAME",
                            "password_env": "ONEC_PASSWORD",
                        },
                        "safety": {
                            "dry_run_env": "DRY_RUN",
                            "allow_document_write_env": "ALLOW_DOCUMENT_WRITE",
                            "allow_document_posting_env": "ALLOW_DOCUMENT_POSTING",
                            "integration_comment": "test",
                        },
                        "metadata": {
                            "stock_register": "",
                            "stock_dimensions": {},
                            "stock_resources": {},
                            "documents": {},
                        },
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            config = load_config(path)
        service = OneCIntegrationService(config, FakeClient())

        with self.assertRaises(IntegrationError) as ctx:
            service.stocks()

        self.assertEqual(ctx.exception.code, ErrorCode.VALIDATION_ERROR)

    def test_stock_query_uses_known_metadata_names(self):
        config = load_config()
        service = OneCIntegrationService(config, FakeClient())

        response = service.stocks()

        self.assertTrue(response["success"])
        self.assertEqual(response["items"][0]["available"], 7)


if __name__ == "__main__":
    unittest.main()
