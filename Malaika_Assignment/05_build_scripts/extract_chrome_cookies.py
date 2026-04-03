"""
Extracts Google/NotebookLM cookies from Chrome and saves as notebooklm-py storage_state.json
"""
import os, json, shutil, sqlite3, base64, re
from pathlib import Path
from datetime import datetime

# Chrome paths
CHROME_DATA = Path(os.environ['LOCALAPPDATA']) / 'Google/Chrome/User Data/Default'
COOKIES_DB   = CHROME_DATA / 'Network/Cookies'
LOCAL_STATE  = Path(os.environ['LOCALAPPDATA']) / 'Google/Chrome/User Data/Local State'
TEMP_COOKIES = Path(__file__).parent / 'temp_cookies.db'

# Output path notebooklm-py expects
OUTPUT_DIR = Path.home() / '.notebooklm/profiles/default'
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_FILE = OUTPUT_DIR / 'storage_state.json'

# Also write legacy path
LEGACY_DIR = Path.home() / '.notebooklm'
LEGACY_FILE = LEGACY_DIR / 'storage_state.json'

def get_encryption_key():
    with open(LOCAL_STATE, 'r', encoding='utf-8') as f:
        local_state = json.load(f)
    encrypted_key = base64.b64decode(local_state['os_crypt']['encrypted_key'])
    encrypted_key = encrypted_key[5:]  # Remove DPAPI prefix
    import win32crypt
    return win32crypt.CryptUnprotectData(encrypted_key, None, None, None, 0)[1]

def decrypt_cookie(encrypted_value, key):
    try:
        from Crypto.Cipher import AES
        if encrypted_value[:3] == b'v10' or encrypted_value[:3] == b'v11':
            nonce = encrypted_value[3:15]
            ciphertext = encrypted_value[15:-16]
            tag = encrypted_value[-16:]
            cipher = AES.new(key, AES.MODE_GCM, nonce=nonce)
            return cipher.decrypt_and_verify(ciphertext, tag).decode('utf-8')
    except Exception as e:
        pass
    try:
        import win32crypt
        return win32crypt.CryptUnprotectData(encrypted_value, None, None, None, 0)[1].decode('utf-8')
    except:
        return ''

# Copy DB (Chrome locks it)
shutil.copy2(COOKIES_DB, TEMP_COOKIES)

key = get_encryption_key()

conn = sqlite3.connect(TEMP_COOKIES)
cursor = conn.cursor()

# Get all Google cookies
cursor.execute("""
    SELECT host_key, name, encrypted_value, path, expires_utc, is_secure, is_httponly, samesite
    FROM cookies
    WHERE host_key LIKE '%google.com%' OR host_key LIKE '%notebooklm%'
    ORDER BY host_key, name
""")

cookies = []
target_names = {'SID','HSID','SSID','APISID','SAPISID','__Secure-1PSID','__Secure-3PSID',
                '__Secure-1PAPISID','__Secure-3PAPISID','__Secure-1PSIDTS','__Secure-3PSIDTS',
                '__Secure-1PSIDCC','__Secure-3PSIDCC','NID','1P_JAR','CONSENT'}

found_names = set()
for host, name, enc_val, path, expires, secure, httponly, samesite in cursor.fetchall():
    value = decrypt_cookie(enc_val, key)
    if value:
        cookies.append({
            "name": name,
            "value": value,
            "domain": host,
            "path": path,
            "expires": expires / 1000000 - 11644473600 if expires > 0 else -1,
            "httpOnly": bool(httponly),
            "secure": bool(secure),
            "sameSite": ["Unspecified","NoRestriction","Lax","Strict"][samesite] if samesite < 4 else "Unspecified"
        })
        if name in target_names:
            found_names.add(name)

conn.close()
TEMP_COOKIES.unlink()

storage_state = {
    "cookies": cookies,
    "origins": []
}

with open(OUTPUT_FILE, 'w') as f:
    json.dump(storage_state, f, indent=2)

with open(LEGACY_FILE, 'w') as f:
    json.dump(storage_state, f, indent=2)

print(f"Extracted {len(cookies)} Google cookies")
print(f"Key cookies found: {found_names}")
print(f"Saved to: {OUTPUT_FILE}")
print(f"Also saved to: {LEGACY_FILE}")
