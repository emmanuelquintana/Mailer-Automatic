import unittest

from enviar_correos import bulk_update_contact_status, update_contact_status


class ContactStatusTests(unittest.TestCase):
    def test_update_contact_status_changes_selected_row(self) -> None:
        rows = [
            {"estado": "Pendiente"},
            {"estado": "Enviado"},
        ]

        update_contact_status(rows, 0, "Revisar")

        self.assertEqual(rows[0]["estado"], "Revisar")
        self.assertEqual(rows[1]["estado"], "Enviado")

    def test_bulk_update_contact_status_changes_all_rows(self) -> None:
        rows = [
            {"estado": "Pendiente"},
            {"estado": "Error"},
            {"estado": ""},
        ]

        bulk_update_contact_status(rows, "Pendiente")

        for row in rows:
            self.assertEqual(row["estado"], "Pendiente")


if __name__ == "__main__":
    unittest.main()
