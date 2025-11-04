import base64

imagenes = [
    "superior_izquierda.jpg",
    "superior_derecha.jpg",
    "centro.jpg",
    "abajo.jpg"
]

for img in imagenes:
    with open(img, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    print(f"{img}:")
    print(b64)  # SIN TRUNCAR
    print("\n" + "="*50 + "\n")