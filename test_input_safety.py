import csv
import io
import os
import tempfile
import unittest
from datetime import timedelta

import app


def legacy_rows(text):
    """変更前の read_file と同じ抽出（検証なし）。"""
    data = []
    reader = csv.reader(io.StringIO(text))
    next(reader)
    for row in reader:
        if row[1] != "":
            data.append(row)
    return data


def sample_row(start="09:00", end="18:00", break_start="12:00", break_end="13:00"):
    row = [""] * 16
    row[0] = "2024-01-01"
    row[1] = start
    row[2] = break_start
    row[3] = break_end
    row[6] = end
    row[15] = "8:00"
    return row


def write_csv(directory, name, text):
    path = os.path.join(directory, name)
    with open(path, "w", encoding="utf-8", newline="") as handle:
        handle.write(text)
    return path


class InputSafetyTests(unittest.TestCase):
    def test_crlf_matches_previous_reader(self):
        header = ",".join(f"c{i}" for i in range(16))
        body = ",".join(sample_row())
        skipped = ",".join(sample_row(start=""))
        text = header + "\r\n" + body + "\r\n" + skipped + "\r\n"
        with tempfile.TemporaryDirectory() as directory:
            path = write_csv(directory, "attendance.csv", text)
            loaded = app.load_attendance_csv(path)
        self.assertEqual(loaded, legacy_rows(text.replace("\r\n", "\n")))

    def test_blank_attendance_short_row_is_skipped(self):
        header = ",".join(f"c{i}" for i in range(16))
        text = header + "\n" + "date,,\n" + ",".join(sample_row()) + "\n"
        with tempfile.TemporaryDirectory() as directory:
            path = write_csv(directory, "attendance.csv", text)
            loaded = app.load_attendance_csv(path)
        self.assertEqual(len(loaded), 1)
        self.assertEqual(loaded[0][1], "09:00")

    def test_short_attended_row_is_rejected(self):
        text = "h1,h2\n2024-01-01,09:00\n"
        with tempfile.TemporaryDirectory() as directory:
            path = write_csv(directory, "short.csv", text)
            with self.assertRaises(app.CsvInputError):
                app.load_attendance_csv(path)

    def test_null_byte_rejected(self):
        header = ",".join(f"c{i}" for i in range(16))
        row = sample_row()
        row[0] = "bad\x00date"
        text = header + "\n" + ",".join(row) + "\n"
        with tempfile.TemporaryDirectory() as directory:
            path = write_csv(directory, "null.csv", text)
            with self.assertRaises(app.CsvInputError):
                app.load_attendance_csv(path)

    def test_directory_and_device_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaises(app.CsvInputError):
                app.load_attendance_csv(directory)
        with self.assertRaises(app.CsvInputError):
            app.load_attendance_csv("/dev/null")
        with self.assertRaises(app.CsvInputError):
            app.load_attendance_csv("/proc/self/environ")

    def test_oversize_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            path = os.path.join(directory, "big.csv")
            with open(path, "wb") as handle:
                handle.write(b"a" * (app.MAX_CSV_BYTES + 1))
            with self.assertRaises(app.CsvInputError):
                app.load_attendance_csv(path)

    def test_drop_list_does_not_execute_commands(self):
        root = app.WorkTimeApp()
        try:
            root.withdraw()
            paths = root._dropped_paths("{/tmp/my file.csv}")
            self.assertEqual(paths, ["/tmp/my file.csv"])
            windows = root._dropped_paths("{C:\\Users\\foo\\bar.csv}")
            self.assertEqual(windows, ["C:\\Users\\foo\\bar.csv"])
            multiple = root._dropped_paths("{/tmp/a.csv} {/tmp/b.csv}")
            self.assertEqual(multiple, ["/tmp/a.csv", "/tmp/b.csv"])
            literal = root._dropped_paths("[exec echo pwned]")
            self.assertEqual(literal, ["[exec", "echo", "pwned]"])
            with self.assertRaises(app.CsvInputError):
                root._dropped_paths("{/tmp/a.csv}; rm -rf /")
            with self.assertRaises(app.CsvInputError):
                root._dropped_paths(None)
        finally:
            root.destroy()

    def test_process_file_total_unchanged(self):
        header = ",".join(f"c{i}" for i in range(16))
        text = header + "\n" + ",".join(sample_row()) + "\n"
        root = app.WorkTimeApp()
        try:
            root.withdraw()
            with tempfile.TemporaryDirectory() as directory:
                path = write_csv(directory, "attendance.csv", text)
                hours, minutes, data = root.process_file(path)
            self.assertEqual((hours, minutes), (8, 0))
            self.assertEqual(data[0][1], "09:00")
            self.assertEqual(data[0][6], "18:00")
        finally:
            root.destroy()

    def test_symlink_to_csv_allowed_and_device_symlink_rejected(self):
        header = ",".join(f"c{i}" for i in range(16))
        text = header + "\n" + ",".join(sample_row()) + "\n"
        with tempfile.TemporaryDirectory() as directory:
            target = write_csv(directory, "attendance.csv", text)
            link = os.path.join(directory, "link.csv")
            os.symlink(target, link)
            self.assertEqual(len(app.load_attendance_csv(link)), 1)
            device_link = os.path.join(directory, "devnull.csv")
            os.symlink("/dev/null", device_link)
            with self.assertRaises(app.CsvInputError):
                app.load_attendance_csv(device_link)

    def test_rounding_unchanged(self):
        self.assertEqual(app.WorkTimeApp.round_time(None, "8:20", "8:30", "start"), "8:30")
        self.assertEqual(app.WorkTimeApp.round_time(None, "8:00", "8:30", "start"), "8:00")
        self.assertEqual(app.WorkTimeApp.round_time(None, "9:10", "そのまま", "start"), "9:10")
        self.assertEqual(app.WorkTimeApp.round_time(None, "-", "9:00", "start"), "-")
        self.assertEqual(app.WorkTimeApp.round_break_start_time(None, "11:50", "12:00"), "12:00")
        self.assertEqual(app.WorkTimeApp.round_break_start_time(None, "11:00", "12:00"), "11:00")
        self.assertEqual(app.WorkTimeApp.round_break_end_time(None, "14:10", "14:25"), "14:25")
        self.assertEqual(app.calculate_break_time("", "13:00"), timedelta(0))
        self.assertEqual(app.calculate_break_time("12:00", "13:00"), timedelta(hours=1))


if __name__ == "__main__":
    unittest.main()
