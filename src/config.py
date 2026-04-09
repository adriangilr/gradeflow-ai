from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv


load_dotenv()


@dataclass(frozen=True)
class Settings:
    credentials_path: Path
    token_path: Path
    download_root: Path
    export_root: Path
    log_root: Path
    default_export_google_doc: str
    default_export_google_sheet: str
    default_export_google_slide: str


def get_settings() -> Settings:
    """
    Centraliza la configuración del proyecto.
    La idea es que cualquier cambio de rutas o formatos pase por aquí.
    """
    return Settings(
        credentials_path=Path(
            os.getenv("GOOGLE_CREDENTIALS_PATH", "credentials/client_secret.json")
        ),
        token_path=Path(os.getenv("GOOGLE_TOKEN_PATH", "token.json")),
        download_root=Path(os.getenv("DOWNLOAD_ROOT", "data/raw")),
        export_root=Path(os.getenv("EXPORT_ROOT", "data/exports")),
        log_root=Path(os.getenv("LOG_ROOT", "data/logs")),
        default_export_google_doc=os.getenv("DEFAULT_EXPORT_GOOGLE_DOC", "pdf"),
        default_export_google_sheet=os.getenv("DEFAULT_EXPORT_GOOGLE_SHEET", "xlsx"),
        default_export_google_slide=os.getenv("DEFAULT_EXPORT_GOOGLE_SLIDE", "pdf"),
    )


def ensure_directories(settings: Settings) -> None:
    """
    Crea las carpetas mínimas para evitar errores al guardar archivos.
    """
    settings.download_root.mkdir(parents=True, exist_ok=True)
    settings.export_root.mkdir(parents=True, exist_ok=True)
    settings.log_root.mkdir(parents=True, exist_ok=True)
