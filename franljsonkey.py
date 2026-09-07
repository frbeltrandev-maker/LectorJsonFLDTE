import hashlib

SECRET_KEY = "348fjeqposaui%//FpouLSDFGkc7513i-*/TRWERa"
machine_id = input("Pega el ID del equipo: ").strip()

clave = hashlib.sha256(
    (machine_id + SECRET_KEY).encode()
).hexdigest()

print("\nCLAVE DE ACTIVACIÓN:")
print(clave)