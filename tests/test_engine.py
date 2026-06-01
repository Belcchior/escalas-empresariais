from datetime import date
import unittest

from engine import gerar_escala_semana, horas_entre, semana_datas


class EngineTests(unittest.TestCase):
    def test_semana_datas_returns_monday_to_sunday(self):
        monday = date(2026, 6, 1)

        result = semana_datas(monday)

        self.assertEqual(len(result), 7)
        self.assertEqual(result[0], date(2026, 6, 1))
        self.assertEqual(result[-1], date(2026, 6, 7))

    def test_horas_entre_calculates_same_day_duration(self):
        self.assertEqual(horas_entre("09:00", "17:30"), 8.5)

    def test_gerar_escala_marks_open_position_without_available_employee(self):
        monday = date(2026, 6, 1)
        estado = {
            "horario": {
                "Segunda": [("09:00", "17:00")],
            },
            "necessidades": {
                "Segunda": {
                    "09:00–17:00": {"caixa": 1},
                },
            },
            "employees": {},
            "rules": {
                "max_horas_dia": 10.0,
                "max_horas_semana": 40.0,
                "min_descanso_horas": 11.0,
            },
        }

        escala = gerar_escala_semana(estado, monday)

        turno = next(iter(escala[monday]))
        self.assertEqual(escala[monday][turno], ["VAGA ABERTA (...caixa)"])


if __name__ == "__main__":
    unittest.main()
