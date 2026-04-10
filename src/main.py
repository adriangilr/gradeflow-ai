from src.auth import get_credentials
from src.config import ensure_directories, get_settings


def main() -> None:
    settings = get_settings()
    ensure_directories(settings)

    creds = get_credentials(
        credentials_path=settings.credentials_path,
        token_path=settings.token_path,
    )

    print("Autenticación correcta.")
    print(f"Token válido: {creds.valid}")


if __name__ == "__main__":
    main()