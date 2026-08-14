import unittest
import json
import os
import tempfile
from unittest.mock import patch

import P2P_Registry_Sync as sync


class PoDescriptionLookupTests(unittest.TestCase):
    def test_matches_complete_invoice_token(self):
        self.assertTrue(
            sync._description_contains_invoice(
                "C/B to Unit TH 629. Invoice no. 1800049318", "1800049318"
            )
        )
        self.assertFalse(
            sync._description_contains_invoice("Invoice 18000493180", "1800049318")
        )

    def test_returns_unique_po_match(self):
        records = [
            {"PO": 116869, "PODesc": "Work completed. Inv no. 1800049318"},
            {"PO": 116870, "Description": "A different invoice"},
        ]

        result = sync._lookup_po_from_description_records("1800049318", records)

        self.assertEqual(result, (116869, "PODesc", 0))

    def test_matches_captured_yardi_po_response(self):
        records = [
            {
                "PO": 116223.0,
                "POCode": "116223",
                "PODesc": (
                    "new swimming pool brush and outdoor fountain chemical "
                    "supply. Inv no. 2607105-09"
                ),
            }
        ]

        result = sync._lookup_po_from_description_records("2607105-09", records)

        self.assertEqual(result, (116223, "PODesc", 0))

    def test_rejects_ambiguous_po_matches(self):
        records = [
            {"PO": 116869, "Description": "Invoice 1800049318"},
            {"PO": 116870, "PODescription": "Correction for 1800049318"},
        ]

        po_id, field, ambiguous_count = (
            sync._lookup_po_from_description_records("1800049318", records)
        )

        self.assertIsNone(po_id)
        self.assertIsNone(field)
        self.assertEqual(ambiguous_count, 2)

    @patch("P2P_Registry_Sync.datetime")
    def test_http_search_matches_website_request(self, mock_datetime):
        mock_datetime.now.return_value.year = 2026

        class Response:
            @staticmethod
            def json():
                return [{
                    "PO": 116223.0,
                    "PODesc": "Inv no. 2607105-09",
                }]

        class Request:
            def __init__(self):
                self.call = None

            def post(self, url, **kwargs):
                self.call = (url, kwargs)
                return Response()

        class Page:
            request = Request()

        page = Page()
        result = sync.search_po_by_description("2607105-09", page, {"Auth": "x"})

        self.assertEqual(result, (116223, "PODesc", 0))
        _, kwargs = page.request.call
        self.assertEqual(kwargs["params"], {
            "dateFrom": "01/01/2026",
            "dateTo": "12/31/2026",
            "search": "2607105-09",
        })
        self.assertEqual(kwargs["data"], "{}")


class MilestoneOrderingTests(unittest.TestCase):
    def test_preserves_timestamp_when_milestones_share_a_date(self):
        milestones = [
            {"Title": "PO Creation", "DtCompleted": "2026-07-23T09:00:00"},
            {"Title": "Approved", "DtCompleted": "2026-07-23T15:30:00"},
        ]

        self.assertEqual(
            sync._current_milestone(milestones, "DtCompleted"),
            ("Approved", "23/7/2026"),
        )

    def test_po_118466_same_date_uses_completed_workflow_priority(self):
        # Model an API response returned newest-first with date-only values
        # and no sequence field. The old index tie-break selected PO Creation.
        milestones = [
            {"Title": "Approved", "DtCompleted": "2026-07-23"},
            {"Title": "Workflow: Condo Manager", "DtCompleted": "2026-07-23"},
            {"Title": "Workflow: Create PO", "DtCompleted": "2026-07-23"},
            {"Title": "PO Creation", "DtCompleted": "2026-07-23"},
        ]

        self.assertEqual(
            sync._current_milestone(milestones, "DtCompleted"),
            ("Approved", "23/7/2026"),
        )

    def test_recognizes_alternate_sequence_field(self):
        milestones = [
            {"Title": "PO Creation", "DtCompleted": "2026-07-23", "Sequence": 1},
            {"Title": "Approved", "DtCompleted": "2026-07-23", "Sequence": 4},
        ]

        self.assertEqual(
            sync._current_milestone(milestones, "DtCompleted"),
            ("Approved", "23/7/2026"),
        )

    def test_completed_paid_overrides_stale_current_posted_for_ir(self):
        milestones = [
            {
                "Title": "Posted",
                "RevisedDate": "2026-08-04",
                "IsCurrent": True,
            },
            {
                "Title": "Paid",
                "RevisedDate": "2026-08-04",
                "IsComplete": True,
                "CheckNum": "7590",
            },
        ]

        class Response:
            @staticmethod
            def json():
                return milestones

        class Request:
            @staticmethod
            def post(*args, **kwargs):
                return Response()

        class Page:
            request = Request()

        self.assertEqual(
            sync.lookup_ir_milestone(145011, Page(), {}),
            ("IR: Cheque has issued on 4/8/2026", "7590", "4/8/2026"),
        )

    def test_existing_cheque_date_produces_terminal_status(self):
        self.assertEqual(
            sync.issued_cheque_status("2026-08-04 00:00:00"),
            "IR: Cheque has issued on 4/8/2026",
        )


class AuthenticationStateTests(unittest.TestCase):
    def test_restores_session_cookie_from_saved_state(self):
        state = {
            "cookies": [{
                "name": ".JWTAUTH",
                "value": "header.payload.signature",
                "domain": "example.test",
                "path": "/",
                "expires": -1,
                "httpOnly": True,
                "secure": True,
                "sameSite": "Lax",
            }],
            "origins": [],
        }

        class Context:
            restored_cookies = None

            def add_cookies(self, cookies):
                self.restored_cookies = cookies

        with tempfile.TemporaryDirectory() as temp_dir:
            state_path = os.path.join(temp_dir, "auth_state.json")
            with open(state_path, "w", encoding="utf-8") as state_file:
                json.dump(state, state_file)

            context = Context()
            with patch.object(sync, "AUTH_STATE_PATH", state_path):
                restored = sync.restore_auth_state(context)

        self.assertTrue(restored)
        self.assertEqual(context.restored_cookies[0]["name"], ".JWTAUTH")
        self.assertEqual(context.restored_cookies[0]["expires"], -1)

    def test_saves_state_to_stable_app_data_path(self):
        class Context:
            saved_path = None

            def storage_state(self, path):
                self.saved_path = path

        with tempfile.TemporaryDirectory() as temp_dir:
            state_path = os.path.join(temp_dir, "auth_state.json")
            context = Context()
            with (
                patch.object(sync, "APP_DATA_DIR", temp_dir),
                patch.object(sync, "AUTH_STATE_PATH", state_path),
            ):
                saved = sync.save_auth_state(context)

        self.assertTrue(saved)
        self.assertEqual(context.saved_path, state_path)


if __name__ == "__main__":
    unittest.main()
