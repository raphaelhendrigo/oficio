DEVEXPRESS_LOADING_SELECTORS = [
    ".dxlpLoadingPanel",
    ".dxgvLoadingPanel",
    ".dx-loading-panel",
    ".dx-loading",
]


def dx_input_selectors(base_id: str) -> list[str]:
    if not base_id:
        return []
    base = base_id.lstrip("#")
    return [
        f"#{base}_I",
        f"#{base} input[id$='_I']",
        f"input[name*='{base}']",
    ]


def dx_button_selectors(base_id: str) -> list[str]:
    if not base_id:
        return []
    base = base_id.lstrip("#")
    return [
        f"#{base}_I",
        f"#{base}",
        f"input[name*='{base}']",
        f"button[id*='{base}']",
    ]


DEFAULT_SELECTORS = {
    "login_user": [
        "#ctl00_cphMain_txtUsuario_I",
        "input[name*='Usuario' i]",
        "input[placeholder*='Usu' i]",
    ],
    "login_pass": [
        "#ctl00_cphMain_txtSenha_I",
        "input[type='password']",
    ],
    "login_button": [
        "#ctl00_cphMain_btnLogin_I",
        "input[type='submit'][value*='Entrar' i]",
        "button:has-text('Entrar')",
    ],
    "grid_apopen": [
        "#gvProcesso",
        "#sptMesaTrabalho_gvProcesso",
        "table[id*='gvProcesso']",
    ],
    "loading_overlay": DEVEXPRESS_LOADING_SELECTORS,
}
