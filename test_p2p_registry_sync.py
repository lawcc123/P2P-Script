import unittest
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


if __name__ == "__main__":
    unittest.main()
