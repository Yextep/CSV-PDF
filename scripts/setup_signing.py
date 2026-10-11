"""Genera una firma persistente fuera del repo y configura GitHub Actions."""
import base64
import os
from pathlib import Path
import secrets
import subprocess

directory = Path('/root/.local/share/luma/signing')
directory.mkdir(parents=True, exist_ok=True)
directory.chmod(0o700)
password = directory / 'password'
key = directory / 'luma.p12'
if not password.exists():
    descriptor = os.open(password, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
    with os.fdopen(descriptor, 'w') as output:
        output.write(secrets.token_urlsafe(40))
if not key.exists():
    subprocess.run(['keytool', '-genkeypair', '-keystore', str(key), '-storetype', 'PKCS12',
        '-storepass:file', str(password), '-keypass:file', str(password), '-alias', 'luma',
        '-keyalg', 'RSA', '-keysize', '3072', '-validity', '10000',
        '-dname', 'CN=Luma, OU=Android, O=Luma', '-noprompt'], check=True)
    key.chmod(0o600)
for name, value in [('LUMA_SIGNING_KEYSTORE', base64.b64encode(key.read_bytes())), ('LUMA_SIGNING_PASSWORD', password.read_bytes())]:
    subprocess.run(['gh', 'secret', 'set', name, '--repo', 'Yextep/CSV-PDF'], input=value, check=True)
print('Firma de Luma guardada fuera del código y configurada en GitHub Actions.')
