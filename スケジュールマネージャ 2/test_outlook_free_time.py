import json
import threading
import unittest
from datetime import date, datetime
from unittest.mock import Mock, patch

from outlook_free_time import (
    AppError, common_free_periods, get_views_from_outlook, validate_search,
    views_for_day,
)


class SearchLogicTests(unittest.TestCase):
    def test_validates_and_deduplicates_inputs(self):
        result = validate_search(
            [" A@example.com ", "a@example.com", "b@example.com", ""],
            "2026-10-07", "2026-10-08", "09:00", "18:00")
        self.assertEqual(result[0], ["A@example.com", "b@example.com"])
        self.assertEqual(result[1:3], (date(2026, 10, 7), date(2026, 10, 8)))
        with self.assertRaises(AppError):
            validate_search(["a@example.com"], "2026-10-08", "2026-10-07", "09:00", "18:00")
        with self.assertRaises(AppError):
            validate_search(["a@example.com"], "2026-10-07", "2026-10-07", "09:15", "18:00")
        self.assertEqual(validate_search(["a@example.com"], "2026-10-07", "2026-10-07", "00:00", "24:00")[-1], 1440)

    def test_intersects_everyone_and_excludes_past_slots(self):
        day = date(2026, 10, 7)
        now = datetime(2026, 10, 7, 9, 10)
        periods = common_free_periods(["0002", "2000"], day, 9 * 60, 11 * 60, now)
        self.assertEqual([(a.strftime("%H:%M"), b.strftime("%H:%M")) for a, b in periods],
                         [("09:30", "10:30")])

    def test_rejects_incomplete_availability(self):
        with self.assertRaises(AppError):
            common_free_periods(["00", "0"], date(2026, 10, 7), 540, 600,
                                datetime(2026, 10, 6))

    @patch("outlook_free_time.subprocess.Popen")
    def test_outlook_includes_self_and_requires_every_person(self, popen):
        process = Mock()
        process.poll.return_value = 0
        popen.return_value = process
        response = {"ok": True, "self_view": "1" * 96, "values": [
            {"email": "b@example.com", "view": "0" * 96},
            {"email": "a@example.com", "view": "2" * 96},
        ]}
        process.communicate.return_value = (json.dumps(response), "")
        views = get_views_from_outlook(["a@example.com", "b@example.com"],
                                       date(2026, 10, 7), date(2026, 10, 8), threading.Event())
        self.assertEqual(views, ["1" * 96, "2" * 96, "0" * 96])
        self.assertEqual(views_for_day(views, date(2026, 10, 7), date(2026, 10, 8),
                                       540, 600), ["11", "22", "00"])
        response["values"].pop(0)
        process.communicate.return_value = (json.dumps(response), "")
        with self.assertRaises(AppError):
            get_views_from_outlook(["a@example.com", "b@example.com"],
                                   date(2026, 10, 7), date(2026, 10, 8), threading.Event())
        response["values"] = [{"email": "a@example.com", "view": "0" * 96},
                              {"email": "b@example.com", "view": "0" * 96}]
        response.pop("self_view")
        process.communicate.return_value = (json.dumps(response), "")
        with self.assertRaisesRegex(AppError, "自分の予定"):
            get_views_from_outlook(["a@example.com", "b@example.com"],
                                   date(2026, 10, 7), date(2026, 10, 8), threading.Event())

    @patch("outlook_free_time.subprocess.Popen")
    def test_self_busy_slot_is_excluded_even_when_others_are_free(self, popen):
        process = Mock()
        process.poll.return_value = 0
        process.communicate.return_value = (json.dumps({
            "ok": True,
            "self_view": "0" * 19 + "2" + "0" * 28,
            "values": [{"email": "a@example.com", "view": "0" * 48}],
        }), "")
        popen.return_value = process
        day = date(2026, 10, 7)
        views = get_views_from_outlook(["a@example.com"], day, day, threading.Event())
        slots = views_for_day(views, day, day, 540, 600)
        periods = common_free_periods(slots, day, 540, 600, datetime(2026, 10, 6))
        self.assertEqual([(start.strftime("%H:%M"), end.strftime("%H:%M"))
                          for start, end in periods], [("09:00", "09:30")])

    @patch("outlook_free_time.subprocess.Popen")
    def test_outlook_error_is_shown(self, popen):
        process = Mock()
        process.poll.return_value = 1
        process.communicate.return_value = ('{"ok":false,"error":"Outlook unavailable"}', "")
        popen.return_value = process
        with self.assertRaisesRegex(AppError, "Outlook unavailable"):
            get_views_from_outlook(["a@example.com"], date(2026, 10, 7),
                                   date(2026, 10, 7), threading.Event())


if __name__ == "__main__":
    unittest.main()
