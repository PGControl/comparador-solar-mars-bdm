import sys

with open('FormComparador.vb', 'r', encoding='utf-8', errors='ignore') as f:
    text = f.read()

replacements = {
    "BOTONES DE SELECCIÃƒâ€œN": "BOTONES DE SELECCIÓN",
    "BOTONES DE SELECCIA'A,\"N": "BOTONES DE SELECCIÓN",
    "LA'A,\"GICA": "LÓGICA",
    "LA\"GICA": "LÓGICA",
    "COMPARACIA'A,\"N": "COMPARACIÓN",
    "elegA'A,A": "elegí",
    "AtenciA'A,A3n": "Atención",
    "ComparaciA3n": "Comparación",
    "DescripciA3n": "Descripción",
    "DirecciA3n": "Dirección",
    "direcciA3n": "dirección",
    "ADesea": "¿Desea",
    "ConfirmaciA3n": "Confirmación",
    "A%xito": "Éxito",
    "Ã‰xito": "Éxito",
    "Ãƒâ€°xito": "Éxito",
    "A'A,Axito": "Éxito",
    "automAtico": "automático",
    "BAsqueda": "Búsqueda",
    "AcciA3n": "Acción",
    "DescripciA'A,A3n": "Descripción",
    "DirecciA'A,A3n": "Dirección",
    "DescripciA'A,A3n": "Descripción",
    "DirecciA'A,A3n": "Dirección",
    "estA": "está",
    "automAticamente": "automáticamente",
    "InformaciA3n": "Información",
    "InformaciA'A,A3n": "Información",
    "InformaciA'A,A3n": "Información",
    "vA'A,Alida": "válida",
    "vA'A,Alida": "válida",
    "ÃƒÂ³": "ó",
    "Ã‚¿": "¿",
    "ÃƒÂ¡": "á",
    "ÃƒÂº": "ú",
    "A'A,Axito": "Éxito",
    "LA'A,\"GICA": "LÓGICA",
    "SELECCIA'A,\"N": "SELECCIÓN",
    "COMPARACIA'A,\"N": "COMPARACIÓN",
    "AcciA'A,A3n": "Acción",
}

for old, new in replacements.items():
    text = text.replace(old, new)

with open('FormComparador.vb', 'w', encoding='utf-8') as f:
    f.write(text)
