from __future__ import annotations

import main as fluxo_atos  # type: ignore


AUTHORIZED = "TC/007902/2022,TC/008636/2022,TC/008084/2023,TC/013838/2023,TC/018149/2024"


def _enable_prod_cleanup(monkeypatch) -> None:
    monkeypatch.setenv("ENVIRONMENT", "producao")
    monkeypatch.setenv("SAFE_DELETE_OWN_DRAFTS", "true")
    monkeypatch.setenv("RUN_PROD_DESTRUCTIVE_CLEANUP", "true")
    monkeypatch.setenv("ONLY_PROCESSOS_AUTHORIZED", AUTHORIZED)


def test_destructive_cleanup_authorized_only_for_five(monkeypatch) -> None:
    _enable_prod_cleanup(monkeypatch)
    monkeypatch.setenv("FORCE_DELETE_OLD_OFICIO_SSG", "true")

    ok, reason = fluxo_atos.destructive_cleanup_authorized(
        "TC/007902/2022",
        required_force_flag="FORCE_DELETE_OLD_OFICIO_SSG",
    )
    assert ok, reason

    ok, reason = fluxo_atos.destructive_cleanup_authorized(
        "TC/000000/2026",
        required_force_flag="FORCE_DELETE_OLD_OFICIO_SSG",
    )
    assert not ok
    assert "fora" in reason


def test_destructive_cleanup_blocks_without_prod_flags(monkeypatch) -> None:
    monkeypatch.setenv("ENVIRONMENT", "homologacao")
    monkeypatch.setenv("SAFE_DELETE_OWN_DRAFTS", "true")
    monkeypatch.setenv("RUN_PROD_DESTRUCTIVE_CLEANUP", "true")
    monkeypatch.setenv("ONLY_PROCESSOS_AUTHORIZED", AUTHORIZED)
    monkeypatch.setenv("FORCE_DELETE_OLD_OFICIO_SSG", "true")

    ok, reason = fluxo_atos.destructive_cleanup_authorized(
        "TC/007902/2022",
        required_force_flag="FORCE_DELETE_OLD_OFICIO_SSG",
    )
    assert not ok
    assert "ENVIRONMENT" in reason


def test_destructive_cleanup_blocks_missing_specific_force(monkeypatch) -> None:
    _enable_prod_cleanup(monkeypatch)
    monkeypatch.delenv("FORCE_RECREATE_COMUNICACAO", raising=False)

    ok, reason = fluxo_atos.destructive_cleanup_authorized(
        "TC/007902/2022",
        required_force_flag="FORCE_RECREATE_COMUNICACAO",
    )
    assert not ok
    assert "FORCE_RECREATE_COMUNICACAO" in reason


def test_terminal_and_pending_status_helpers() -> None:
    assert fluxo_atos.is_terminal_final_status("Ofício SSG Assinado")
    assert fluxo_atos.is_terminal_final_status("Documento Finalizado")
    assert fluxo_atos.is_terminal_final_status("Publicado")
    assert not fluxo_atos.is_terminal_final_status("Em assinatura")
    assert fluxo_atos.is_pending_signature_status("Em assinatura")
    assert fluxo_atos.is_pending_signature_status("Aguardando assinatura")
    assert fluxo_atos.is_pending_signature_status("Pendente de assinatura")


def test_post_conclusion_policy_skips_signature_and_tramitacao(monkeypatch) -> None:
    monkeypatch.setenv("REQUEST_SIGNATURE", "true")
    monkeypatch.setenv("ASSINANTE_NOME", "Roseli Chaves")
    monkeypatch.setenv("TRAMITAR_DESTINO", "Em assinatura")
    monkeypatch.setenv("SKIP_SIGNATURE", "true")
    monkeypatch.setenv("SKIP_TRAMITACAO", "true")
    monkeypatch.setenv("STOP_AFTER_OFICIO_CONCLUIDO", "true")

    policy = fluxo_atos._post_conclusion_policy()
    assert policy["signature_required"] is False
    assert policy["tramitacao_required"] is False
    assert policy["stop_after_oficio_concluido"] is True


def test_post_conclusion_policy_allows_legacy_when_not_skipped(monkeypatch) -> None:
    monkeypatch.setenv("REQUEST_SIGNATURE", "true")
    monkeypatch.setenv("ASSINANTE_NOME", "Roseli Chaves")
    monkeypatch.setenv("TRAMITAR_DESTINO", "Em assinatura")
    monkeypatch.delenv("SKIP_SIGNATURE", raising=False)
    monkeypatch.delenv("SKIP_TRAMITACAO", raising=False)
    monkeypatch.delenv("STOP_AFTER_OFICIO_CONCLUIDO", raising=False)

    policy = fluxo_atos._post_conclusion_policy()
    assert policy["signature_required"] is True
    assert policy["tramitacao_required"] is True


def test_robot_created_comunicacao_detection_is_conservative() -> None:
    assert fluxo_atos._is_robot_created_comunicacao_text("Oficio UTAP - modelo Educação - gerado automaticamente")
    assert fluxo_atos._is_robot_created_comunicacao_text("robô Euclides")
    assert not fluxo_atos._is_robot_created_comunicacao_text("Secretaria Municipal de Educação Domingos Dissei")
