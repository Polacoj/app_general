def main() -> None:
    """Entry point used by the installed ``app-general`` command."""
    import flet as ft

    from .main import main as pagina_principal

    ft.run(pagina_principal)
