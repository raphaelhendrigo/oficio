from __future__ import annotations

from datetime import date as real_date

from web import runner, store


class _FixedDate:
    @classmethod
    def today(cls) -> real_date:
        return real_date(2026, 5, 29)


def test_default_env_matches_reiteracao_prod_runner(monkeypatch) -> None:
    monkeypatch.setattr(runner, "date", _FixedDate)

    env = runner._default_env_for_tipo("REITERACAO")

    assert env["FORCE_TIPO"] == "REITERACAO"
    assert env["REQUEST_SIGNATURE"] == "true"
    assert env["ASSINANTE_NOME"] == "Roseli Chaves"
    assert env["SKIP_TRAMITACAO"] == "true"
    assert env["OFICIO_REQUIRE_PIECE_NUMBER_IN_ENCAMINHA"] == "true"
    assert env["REITERACAO_REQUIRE_AUTO_EXTRACT"] == "true"
    assert env["DATA_OFICIO"] == "29/05/2026"
    assert env["OFICIO_DATA"] == "29/05/2026"


def test_default_env_keeps_dilacao_piece_number_optional(monkeypatch) -> None:
    monkeypatch.setattr(runner, "date", _FixedDate)

    env = runner._default_env_for_tipo("DILACAO")

    assert env["FORCE_TIPO"] == "DILACAO"
    assert env["OFICIO_REQUIRE_PIECE_NUMBER_IN_ENCAMINHA"] == "false"
    assert "REITERACAO_REQUIRE_AUTO_EXTRACT" not in env


def test_scan_events_only_marks_ok_after_signature(monkeypatch) -> None:
    monkeypatch.setattr(store, "persist", lambda job: None)
    job = store.new_job(
        excel_filename="teste.xlsx",
        processos=[{"processo": "TC/016628/2024", "tipo": "REITERACAO"}],
    )
    try:
        runner._scan_for_events(
            job.id,
            "Iniciando pipeline do processo TC/016628/2024.",
            ["TC/016628/2024"],
        )
        runner._scan_for_events(
            job.id,
            "Ato Oficio SSG concluido para TC/016628/2024.",
            ["TC/016628/2024"],
        )
        assert store.get_job(job.id).processos[0].status == "running"

        runner._scan_for_events(
            job.id,
            "Assinatura solicitada para Roseli Chaves no processo TC/016628/2024.",
            ["TC/016628/2024"],
        )
        assert store.get_job(job.id).processos[0].status == "ok"
    finally:
        with store._lock:
            store._jobs.pop(job.id, None)
