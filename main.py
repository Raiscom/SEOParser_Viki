"""Точка входа приложения SEO API Парсер VIKI."""

from app.config import configure_runtime_environment, get_settings
from app.logging_setup import setup_logging


def main() -> None:
    """Запускает приложение и проверяет загрузку настроек."""
    configure_runtime_environment()
    setup_logging()
    get_settings()
    from app.gui import SeoParserApp

    app = SeoParserApp()
    app.run()


if __name__ == "__main__":
    main()
