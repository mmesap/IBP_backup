string = """
Teleperformance
Lenovo
QMC Telecom
Qualtrics
Superior Energy Services
IBP
Alfa
AWS
Infobip
Meltec
Arris
Redeban
Intel Corporation
Abbot - Lafrancol
Gilead
Kyndryl
Bizagi
IBM
Ingram Micro
Epson Latam
Lablium 
TD SYNNEX
Oracle
Microsoft
Motorola Solutions
SAP
VIVO Mexico
NCR
Procaps
Innovasport
Indices CognoSight
Sales Force
Hassar Mexico
Master Card
"""
EMPRESAS = string.split("\n")

for empresa in EMPRESAS:
    print(empresa.upper().strip().replace("  ", " ").replace(",", "").replace(".", "").replace("(", "").replace(")", ""))