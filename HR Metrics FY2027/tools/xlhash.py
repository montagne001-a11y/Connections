import hashlib,base64,struct,os
def xl_hash(pw, salt_b64, spin=100000):
    salt=base64.b64decode(salt_b64)
    h=hashlib.sha512(salt+pw.encode('utf-16-le')).digest()
    for i in range(spin):
        h=hashlib.sha512(h+struct.pack('<I',i)).digest()
    return base64.b64encode(h).decode()
def new_hash(pw, spin=100000):
    salt=base64.b64encode(os.urandom(16)).decode()
    return salt, xl_hash(pw,salt,spin)
