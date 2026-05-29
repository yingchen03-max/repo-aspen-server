import ipaddress
import os
import socket
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Iterable, List, Tuple

from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.x509.oid import NameOID


def resolve_ssl_paths(cert_file: str, key_file: str, base_dir: Path | None = None) -> Tuple[Path, Path]:
    """Resolve SSL paths, treating relative paths as relative to base_dir."""
    base_dir = base_dir or Path(__file__).parent
    cert_path = Path(cert_file)
    key_path = Path(key_file)

    if not cert_path.is_absolute():
        cert_path = base_dir / cert_path
    if not key_path.is_absolute():
        key_path = base_dir / key_path

    return cert_path.resolve(), key_path.resolve()


def _unique_preserve_order(values: Iterable[str]) -> List[str]:
    result: List[str] = []
    seen = set()
    for value in values:
        if not value or value in seen:
            continue
        seen.add(value)
        result.append(value)
    return result


def discover_certificate_hosts() -> List[str]:
    """Collect hostnames and IPs that should be added to certificate SANs."""
    env_hosts = os.getenv("ASPEN_SSL_HOSTS", "")
    requested_hosts = [item.strip() for item in env_hosts.split(",") if item.strip()]

    host_candidates = [
        "localhost",
        "127.0.0.1",
        "::1",
        socket.gethostname(),
        socket.getfqdn(),
    ]

    try:
        addr_info = socket.getaddrinfo(socket.gethostname(), None)
        for item in addr_info:
            host_candidates.append(item[4][0])
    except OSError:
        pass

    return _unique_preserve_order([*requested_hosts, *host_candidates])


def _build_subject_alt_names(hosts: Iterable[str]) -> x509.SubjectAlternativeName:
    san_entries = []
    for host in hosts:
        try:
            san_entries.append(x509.IPAddress(ipaddress.ip_address(host)))
        except ValueError:
            san_entries.append(x509.DNSName(host))
    return x509.SubjectAlternativeName(san_entries)


def _build_subject(common_name: str) -> x509.Name:
    return x509.Name(
        [
            x509.NameAttribute(NameOID.COUNTRY_NAME, os.getenv("ASPEN_SSL_COUNTRY", "CN")),
            x509.NameAttribute(NameOID.ORGANIZATION_NAME, os.getenv("ASPEN_SSL_ORGANIZATION", "Aspen Agent")),
            x509.NameAttribute(NameOID.COMMON_NAME, common_name),
        ]
    )


def generate_self_signed_certificate(cert_path: Path, key_path: Path, hosts: Iterable[str]) -> Tuple[Path, Path]:
    hosts = _unique_preserve_order(hosts)
    if not hosts:
        hosts = ["localhost", "127.0.0.1"]

    cert_path.parent.mkdir(parents=True, exist_ok=True)
    key_path.parent.mkdir(parents=True, exist_ok=True)

    private_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    subject = issuer = _build_subject(hosts[0])
    now = datetime.now(timezone.utc)

    certificate = (
        x509.CertificateBuilder()
        .subject_name(subject)
        .issuer_name(issuer)
        .public_key(private_key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now - timedelta(minutes=5))
        .not_valid_after(now + timedelta(days=825))
        .add_extension(_build_subject_alt_names(hosts), critical=False)
        .add_extension(x509.BasicConstraints(ca=False, path_length=None), critical=True)
        .add_extension(
            x509.KeyUsage(
                digital_signature=True,
                key_encipherment=True,
                content_commitment=False,
                data_encipherment=False,
                key_agreement=False,
                key_cert_sign=False,
                crl_sign=False,
                encipher_only=False,
                decipher_only=False,
            ),
            critical=True,
        )
        .add_extension(x509.ExtendedKeyUsage([x509.oid.ExtendedKeyUsageOID.SERVER_AUTH]), critical=False)
        .sign(private_key, hashes.SHA256())
    )

    key_path.write_bytes(
        private_key.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.TraditionalOpenSSL,
            encryption_algorithm=serialization.NoEncryption(),
        )
    )
    cert_path.write_bytes(certificate.public_bytes(serialization.Encoding.PEM))

    return cert_path, key_path


def ensure_server_certificate(cert_file: str, key_file: str, base_dir: Path | None = None) -> Tuple[Path, Path, bool, List[str]]:
    """
    Ensure a reusable server certificate exists.

    Returns:
        cert_path, key_path, created, hosts
    """
    cert_path, key_path = resolve_ssl_paths(cert_file, key_file, base_dir=base_dir)
    hosts = discover_certificate_hosts()

    if cert_path.exists() and key_path.exists():
        return cert_path, key_path, False, hosts

    generate_self_signed_certificate(cert_path, key_path, hosts)
    return cert_path, key_path, True, hosts
