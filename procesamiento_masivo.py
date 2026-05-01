"""
procesamiento_masivo.py
Lee el Excel maestro y genera todos los convenios en un archivo ZIP.
"""
import pandas as pd
from datetime import datetime
import zipfile, io, os, sys
import importlib.util as _ilu
_BASE = os.path.dirname(os.path.abspath(__file__))
_spec = _ilu.spec_from_file_location("generador", os.path.join(_BASE, "generador.py"))
_gen  = _ilu.module_from_spec(_spec)
_spec.loader.exec_module(_gen)
generar_documento = _gen.generar_documento

def parsear_fecha(texto, obligatoria=True):
    """Parsea texto a date. Si obligatoria=False y esta vacio, retorna None."""
    s = str(texto).strip() if texto is not None else ""
    if s in ("", "nan", "NaT", "None"):
        if not obligatoria:
            return None
        raise ValueError("Fecha requerida pero esta vacia")
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except:
            continue
    raise ValueError(f"Fecha invalida: {texto}")

def nombre_archivo(row):
    est = str(row.get("nombre_estudiante","estudiante")).split()[0].capitalize()
    emp = str(row.get("nombre_empresa","empresa")).split()[0].capitalize()
    fi  = str(row.get("fecha_inicio","")).replace("/","")[:6]
    return f"Convenio_{est}_{emp}_{fi}.docx"

def procesar_excel(ruta_excel: str, carpeta_salida: str = None) -> tuple:
    """
    Lee el Excel y genera un ZIP con todos los convenios.
    Retorna: (zip_bytes, cantidad_ok, lista_errores)
    """
    df = pd.read_excel(ruta_excel, sheet_name="Datos", header=1)
    df.columns = [c.strip() for c in df.columns]

    MAPA = {
        "Tipo de Experiencia":         "tipo_experiencia",
        "Remunerada?":                 "remunerada",
        "Quien paga ARL?":             "quien_paga_arl",
        "Nombre Empresa":              "nombre_empresa",
        "Tipo Sociedad":               "tipo_sociedad",
        "NIT":                         "nit_empresa",
        "Ciudad Empresa":              "ciudad_empresa",
        "Representante Legal":         "nombre_representante",
        "Cedula Rep. Legal":           "cedula_representante",
        "Cargo Rep. Legal":            "cargo_representante",
        "Nombre Estudiante":           "nombre_estudiante",
        "Cedula Estudiante":           "cedula_estudiante",
        "Programa Academico":          "programa_academico",
        "Fecha Inicio (DD/MM/AAAA)":   "fecha_inicio",
        "Fecha Fin (DD/MM/AAAA)":      "fecha_fin",
        "Monto en Letras":             "monto_letras",
        "Monto en Numeros":            "monto_numeros",
        "Nombre Tutor (Empresa)":      "nombre_tutor",
        "Cedula Tutor":                "cedula_tutor",
        "Cargo Tutor":                 "cargo_tutor",
        "Nombre Monitor (EAFIT)":      "nombre_monitor",
        "Cedula Monitor":              "cedula_monitor",
        "Actividad 1":                 "actividad_1",
        "Actividad 2":                 "actividad_2",
        "Actividad 3":                 "actividad_3",
        "Actividad 4":                 "actividad_4",
        "Actividad 5":                 "actividad_5",
        "Actividad 6":                 "actividad_6",
        "Actividad 7":                 "actividad_7",
        "Actividad 8":                 "actividad_8",
        "Fecha Firma (DD/MM/AAAA)":    "fecha_firma",
        "Email Organizacion":          "email_org",
        "Email Estudiante":            "email_est",
        "Email EAFIT":                 "email_eafit",
    }

    # Tambien intentar con tildes (por si el Excel tiene los encabezados con tildes)
    MAPA_TILDE = {
        "¿Remunerada?":                "remunerada",
        "¿Quién paga ARL?":            "quien_paga_arl",
        "Cédula Rep. Legal":           "cedula_representante",
        "Cédula Estudiante":           "cedula_estudiante",
        "Programa Académico":          "programa_academico",
        "Monto en Números":            "monto_numeros",
        "Cédula Tutor":                "cedula_tutor",
        "Cédula Monitor":              "cedula_monitor",
    }
    MAPA.update(MAPA_TILDE)

    df = df.rename(columns=MAPA)
    df = df.dropna(subset=["nombre_estudiante"])

    zip_buffer = io.BytesIO()
    errores = []
    ok = 0

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for idx, row in df.iterrows():
            try:
                fecha_fin_val = parsear_fecha(row.get("fecha_fin",""))
                fecha_firma_val = parsear_fecha(row.get("fecha_firma",""), obligatoria=False) or fecha_fin_val

                datos = {
                    "tipo_experiencia": str(row.get("tipo_experiencia","Practica")).strip(),
                    "remunerada":       str(row.get("remunerada","No")).strip().lower() in ["si","s","yes","1","true","sí"],
                    "quien_paga_arl":   str(row.get("quien_paga_arl","Organizacion")).strip(),
                    "nombre_empresa":   str(row.get("nombre_empresa","")).strip(),
                    "tipo_sociedad":    str(row.get("tipo_sociedad","S.A.S.")).strip(),
                    "nit_empresa":      str(row.get("nit_empresa","")).strip(),
                    "ciudad_empresa":   str(row.get("ciudad_empresa","")).strip(),
                    "nombre_representante": str(row.get("nombre_representante","")).strip(),
                    "cedula_representante": str(row.get("cedula_representante","")).strip(),
                    "cargo_representante":  str(row.get("cargo_representante","")).strip(),
                    "nombre_estudiante":    str(row.get("nombre_estudiante","")).strip(),
                    "cedula_estudiante":    str(row.get("cedula_estudiante","")).strip(),
                    "programa_academico":   str(row.get("programa_academico","")).strip(),
                    "fecha_inicio": parsear_fecha(row.get("fecha_inicio","")),
                    "fecha_fin":    fecha_fin_val,
                    "fecha_firma":  fecha_firma_val,
                    "nombre_tutor":    str(row.get("nombre_tutor","")).strip(),
                    "cedula_tutor":    str(row.get("cedula_tutor","")).strip(),
                    "nombre_monitor":  str(row.get("nombre_monitor","")).strip(),
                    "cedula_monitor":  str(row.get("cedula_monitor","")).strip(),
                    "monto_letras":    str(row.get("monto_letras","") or "").strip(),
                    "monto_numeros":   str(row.get("monto_numeros","") or "").strip(),
                    "actividades": [
                        str(row.get(f"actividad_{i}","") or "").strip()
                        for i in range(1,9)
                        if str(row.get(f"actividad_{i}","") or "").strip()
                    ],
                    "email_org":   str(row.get("email_org","") or "").strip(),
                    "email_est":   str(row.get("email_est","") or "").strip(),
                    "email_eafit": str(row.get("email_eafit","") or "").strip(),
                }
                docx_bytes = generar_documento(datos)
                fname = nombre_archivo(row)
                zf.writestr(fname, docx_bytes)
                ok += 1

                if carpeta_salida:
                    os.makedirs(carpeta_salida, exist_ok=True)
                    with open(os.path.join(carpeta_salida, fname), "wb") as f:
                        f.write(docx_bytes)

            except Exception as e:
                errores.append(f"Fila {idx+2} ({row.get('nombre_estudiante','?')}): {e}")

    zip_buffer.seek(0)
    return zip_buffer.getvalue(), ok, errores


if __name__ == "__main__":
    import time
    ruta = os.path.join(_BASE, "TALENTO_EAFIT_Carga_Masiva.xlsx")
    salida = os.path.join(_BASE, "documentos_generados", "masivo")

    t0 = time.time()
    zip_bytes, ok, errores = procesar_excel(ruta, salida)
    t1 = time.time()

    print(f"\n{'='*50}")
    print(f"Convenios generados: {ok}")
    print(f"Errores:             {len(errores)}")
    print(f"Tiempo total:        {t1-t0:.1f} segundos")
    print(f"ZIP generado:        {len(zip_bytes):,} bytes ({len(zip_bytes)//1024} KB)")
    if errores:
        print("\nErrores:")
        for e in errores:
            print(" ", e)
