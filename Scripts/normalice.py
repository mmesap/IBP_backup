string = """
Automotores Continental
Ecuasanitas
Banco Del Pacífico
Banco General Rumiñahui
Banecuador
Consorcio Nobis
Cooprogreso
Difare
Mega Santamaria Sa
Mega Santamaria S.A.
"""
EMPRESAS = string.split("\n")

for empresa in EMPRESAS:
    print(empresa.upper().strip().replace("  ", " ").replace(",", "").replace(".", "").replace("(", "").replace(")", ""))