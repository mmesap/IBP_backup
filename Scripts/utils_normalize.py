import re
def normalizar_pais(pais):
    if isinstance(pais, str):
        pais = pais.strip().lower()  # Eliminar espacios y convertir a minúsculas
        if re.match(r'^(e\s*e\s*u\s*u|estados\s*unidos)$', pais):
            return 'Estados Unidos'
        elif re.match(r'^puerto\s*rico$', pais):
            return 'Puerto Rico'
        elif re.match(r'^col[ou]mbia$', pais):
            return 'Colombia'
        elif re.match(r'^el\s*salvador$', pais):
            return 'El Salvador'
        elif re.match(r'^nicaragua$', pais):
            return 'Nicaragua'
        elif re.match(r'^honduras$', pais):
            return 'Honduras'
        elif re.match(r'^ecuador$', pais):
            return 'Ecuador'
        elif re.match(r'^costa\s*rica$', pais):
            return 'Costa Rica'
        elif re.match(r'^panam[aá]$', pais):
            return 'Panamá'
        elif re.match(r'^guatemala$', pais):
            return 'Guatemala'
        elif re.match(r'^m[eé]xico$', pais):
            return 'México'
        elif re.match(r'^per[uú]$', pais):
            return 'Perú'
        elif re.match(r'^republica\s*dominicana$', pais):
            return 'República Dominicana'
        elif re.match(r'^espa(ñ|n)a$', pais):
            return 'España'
        elif re.match(r'^br(a|e)sil$', pais):
            return 'Brasil'
        elif re.match(r'^chile$', pais):
            return 'Chile'
        elif re.match(r'^uruguay$', pais):
            return 'Uruguay'
        elif re.match(r'^argentina$', pais):
            return 'Argentina'
        elif re.match(r'^paraguay$', pais):
            return 'Paraguay'
        elif re.match(r'^bolivia$', pais):
            return 'Bolivia'
        elif re.match(r'^venezuela$', pais):
            return 'Venezuela'
        else:
            return pais  # Retornar el valor original si no coincide con ningún patrón
    else:
        return None

def normalizar_rango_empleados(valor):
    valor = str(valor).strip().lower()  # Convertir a minúsculas y eliminar espacios    
    # Expresiones regulares para cada rango permitido
    if re.search(r'^(1[\s\-]*10|1\s*a\s*10)$', valor):
        return '1-10'
    elif re.search(r'^(11[\s\-]*50|11\s*a\s*50)$', valor):
        return '11-50'
    elif re.search(r'^(51[\s\-]*200|50[\s\-]*200|51\s*a\s*200)$', valor):
        return '51-200'
    elif re.search(r'^(201[\s\-]*500|200[\s\-]*500|201\s*a\s*500)$', valor):
        return '201-500'
    elif re.search(r'^(501[\s\-]*1000|500[\s\-]*1000|501\s*a\s*1000)$', valor):
        return '501-1000'
    elif re.search(r'^(1001[\s\-]*5000|1000[\s\-]*5000|1001\s*a\s*5000)$', valor):
        return '1001-5000'
    elif re.search(r'^(5001[\s\-]*10000|5001\s*a\s*10[\.\s]*000|5000[\s\-]*10000|5001\s*a\s*10000)$', valor):
        return '5001-10000'
    elif re.search(r'(10000[\+]*|10[\.\s]*mil|10000)', valor):
        return '10000+'
    else:
        return valor

def normalizar_industria(industria):
    if isinstance(industria, str):
        industria = industria.strip().lower()  # Convertir a minúsculas y eliminar espacios
        return industria.title()
    else:
        return None
    
def normalice_company_name(name):
    if isinstance(name, str):
        name = name.strip().upper()  # Convertir a minúsculas y eliminar espacios
        return name.title()
    else:
        return None