from pathlib import Path

from config import SSL_CERT_FILE, SSL_KEY_FILE
from ssl_utils import ensure_server_certificate


def main():
    cert_path, key_path, created, hosts = ensure_server_certificate(
        SSL_CERT_FILE,
        SSL_KEY_FILE,
        base_dir=Path(__file__).parent,
    )
    status = "created" if created else "reused"
    print(f"SSL certificate {status}: {cert_path}")
    print(f"SSL private key {status}: {key_path}")
    print(f"Certificate SAN: {', '.join(hosts)}")


if __name__ == "__main__":
    main()
