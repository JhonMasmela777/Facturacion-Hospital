
from pathlib import Path
import re
import unicodedata

import pandas as pd
import pdfplumber

from openpyxl.styles import Font, PatternFill, Alignment


# ==================================================
# CONFIGURACION DEL PROYECTO
# ==================================================

CARPETA_ENTRADA = Path("archivos_entrada")
CARPETA_PDFS = CARPETA_ENTRADA / "notas_quirurgicas"

CARPETA_RESULTADOS = Path("resultados")

# Cambiar cuando comience un nuevo mes.
MES_PROCESO = "Noviembre_2026"

CARPETA_RESULTADOS_MES = CARPETA_RESULTADOS / MES_PROCESO

NOMBRE_EXCEL = (
    f"Extracto_Notas_Quirurgicas_{MES_PROCESO}.xlsx"
)

COLUMNAS_EXTRACTO = [
    "ESPECIALISTA",
    "FECHA",
    "IDENTIF",
    "PACIENTE",
    "CODIGO",
    "PROCEDIMIENTOS",
    "ENTIDAD",
    "DIA_CARPETA",
    "ARCHIVO_PDF",
]


# ==================================================
# FUNCIONES GENERALES
# ==================================================

def normalizar_texto(texto):
    """Quita tildes y convierte el texto a minusculas."""

    texto = unicodedata.normalize("NFD", str(texto))

    texto = "".join(
        caracter
        for caracter in texto
        if unicodedata.category(caracter) != "Mn"
    )

    return texto.lower()


def extraer_texto_pdf(ruta_pdf):
    """Lee el texto de todas las paginas del PDF."""

    partes = []

    with pdfplumber.open(ruta_pdf) as pdf:

        for pagina in pdf.pages:

            texto = pagina.extract_text()

            if texto:
                partes.append(texto)

    return "\n".join(partes)


def obtener_archivos_pdf(carpeta):
    """Obtiene todos los PDF, incluyendo subcarpetas."""

    if not carpeta.exists():
        return []

    return sorted(
        (
            archivo
            for archivo in carpeta.rglob("*")
            if archivo.is_file()
            and archivo.suffix.lower() == ".pdf"
        ),
        key=lambda archivo: (
            0,
            int(archivo.stem)
        ) if archivo.stem.isdigit() else (
            1,
            archivo.name.lower()
        )
    )


# ==================================================
# MODULO 1: VERIFICACION DE CARPETAS
# ==================================================

def verificar_estructura():

    print("\n" + "=" * 50)
    print("   AUTOMATIZACION NOTAS QUIRURGICAS")
    print("=" * 50)

    if not CARPETA_PDFS.exists():

        print("\nNo existe la carpeta de entrada:")
        print(CARPETA_PDFS)

        return False

    print("\nCarpeta principal encontrada:", CARPETA_PDFS)

    dias_con_archivos = 0
    total_pdf = 0

    for numero_dia in range(1, 32):

        carpeta_dia = CARPETA_PDFS / f"dia{numero_dia}"

        if not carpeta_dia.exists():
            continue

        archivos = obtener_archivos_pdf(carpeta_dia)

        if archivos:

            dias_con_archivos += 1
            total_pdf += len(archivos)

            print(
                f"Dia {numero_dia}: "
                f"{len(archivos)} PDF encontrados."
            )

    print("\n" + "=" * 50)
    print("   RESUMEN DE CARPETAS")
    print("=" * 50)

    print("Dias con documentos:", dias_con_archivos)
    print("Total de PDF encontrados:", total_pdf)

    if total_pdf == 0:

        print("\nNo se encontraron documentos PDF.")

        return False

    return True


# ==================================================
# MODULO 2: EXTRACCION DE IDENTIFICACION
# ==================================================

def extraer_identificacion(texto):

    patron = re.compile(
        r"\bIdentificaci[oó]n\b\s+"
        r"(?:CC|RC|TI|CE|PA|PEP|PPT|NIT|PT|SI)\s+"
        r"([A-Z0-9]+(?:[./-][A-Z0-9]+)*)",
        re.IGNORECASE
    )

    coincidencia = patron.search(texto)

    if coincidencia:
        return coincidencia.group(1).strip(" .:/-").upper()

    return ""


# ==================================================
# MODULO 3: EXTRACCION DE DATOS GENERALES
# ==================================================

def extraer_datos_generales(texto):

    datos = {
        "ESPECIALISTA": "",
        "FECHA": "",
        "IDENTIF": "",
        "PACIENTE": "",
        "ENTIDAD": "",
    }

    # Limitar encabezado.
    texto_inicial = re.split(
        r"(?im)^\s*DESCRIPCI[ÓO]N\s+(?:DE\s+LA\s+)?CIRUG[IÍ]A\s*:?\s*$",
        texto,
        maxsplit=1
    )[0]

    # ----------------------------------------------
    # PACIENTE
    # ----------------------------------------------

    coincidencia = re.search(
        r"\bPaciente\s*:\s*(.+?)(?=\s+Identificaci[óo]n\b)",
        texto_inicial,
        re.IGNORECASE | re.DOTALL
    )

    if coincidencia:

        datos["PACIENTE"] = re.sub(
            r"\s+",
            " ",
            coincidencia.group(1)
        ).strip()

    # ----------------------------------------------
    # IDENTIFICACION
    # ----------------------------------------------

    # Importante: esta extraccion es independiente
    # de la variable coincidencia utilizada para paciente.

    datos["IDENTIF"] = extraer_identificacion(texto_inicial)

    # ----------------------------------------------
    # ENTIDAD
    # ----------------------------------------------

    coincidencia = re.search(
        r"\bEmpresa\s*:\s*(.+?)(?=\s+(?:Sala|Fecha|Sede)\s*:)",
        texto_inicial,
        re.IGNORECASE | re.DOTALL
    )

    if coincidencia:

        datos["ENTIDAD"] = re.sub(
            r"\s+",
            " ",
            coincidencia.group(1)
        ).strip()

   
    # ----------------------------------------------
    # ESPECIALISTA
    # ----------------------------------------------

    # FORMATO 1:
    # Cirujano: NOMBRE Especialidad: ESPECIALIDAD

    coincidencia = re.search(
        r"\bCirujano\s*:\s*(.+?)"
        r"(?=\s+Especialidad\s*:)",
        texto_inicial,
        re.IGNORECASE | re.DOTALL
    )

    # FORMATO 2:
    # CIRUJANO ES235 NOMBRE
    # Especialidad CIRUGIA PEDIATRICA

    if not coincidencia:

        coincidencia = re.search(
            r"\bCIRUJANO\s+[A-Z]{1,4}\d{2,6}\s+"
            r"(.+?)"
            r"(?=\s+Especialidad\b)",
            texto,
            re.IGNORECASE | re.DOTALL
        )

    if coincidencia:

        datos["ESPECIALISTA"] = re.sub(
            r"\s+",
            " ",
            coincidencia.group(1)
        ).strip()

    # ----------------------------------------------
    # FECHA DEL ACTO QUIRURGICO
    # ----------------------------------------------

        # ----------------------------------------------
    # FECHA DE PROGRAMACION DE CIRUGIA
    # ----------------------------------------------

    coincidencia = re.search(
        r"\bFecha\s+programaci[oó]n\s+de\s+cirug[ií]a\s*:\s*(\d{2}/\d{2}/\d{4})",
        texto,
        re.IGNORECASE
    )

    if coincidencia:
        datos["FECHA"] = coincidencia.group(1).strip()
    else:
        datos["FECHA"] = ""

    return datos

# ==================================================
# MODULO 4: LIMPIEZA DE DESCRIPCIONES
# ==================================================

def limpiar_descripcion_procedimiento(descripcion):
    """Limpia informacion adicional de la descripcion."""

    patrones_corte = [
        r"\bcirujano\s*:",
        r"\bespecialista\s*:",
        r"\bespecialidad\s*:?",
        r"\bv[ií]a\s*:",
        r"\bdescripcion\s+(?:de\s+la\s+)?cirugia\b",
    ]

    for patron in patrones_corte:

        coincidencia = re.search(
            patron,
            descripcion,
            flags=re.IGNORECASE
        )

        if coincidencia:

            descripcion = descripcion[:coincidencia.start()]

    # Quitar identificador de grupo quirurgico al final.
    descripcion = re.sub(
        r"\s+Y\d{2,3}\s*$",
        "",
        descripcion.strip(),
        flags=re.IGNORECASE
    )

    descripcion = re.sub(
        r"\s+",
        " ",
        descripcion
    ).strip()

    return descripcion


# ==================================================
# MODULO 5: EXTRACCION DE PROCEDIMIENTOS
# ==================================================




# ==================================================
# MODULO 5: EXTRACCION INTELIGENTE DE PROCEDIMIENTOS
# ==================================================

def extraer_procedimientos(texto):
    """
    Reconoce los dos formatos de HOSVITAL:

    1. CIRUGIAS + DESCRIPCION CIRUGIA.
    2. DESCRIPCION DE CIRUGIAS sin tabla CIRUGIAS.

    Evita duplicar procedimientos presentes en ambas secciones.
    Conserva codigo, descripcion y cantidad.
    """

    procedimientos_tabla = []
    procedimientos_descripcion = []

    lineas = texto.splitlines()

    # ------------------------------------------------
    # FUNCION PARA RECONOCER ENCABEZADOS
    # ------------------------------------------------

    patron_descripcion = re.compile(
        r"^\s*DESCRIPCI[ÓO]N\s+"
        r"(?:DE\s+(?:LA\s+)?)?"
        r"CIRUG[IÍ]AS?\b",
        re.IGNORECASE
    )

    patron_cirugias = re.compile(
        r"^\s*CIRUG[IÍ]AS\s*$",
        re.IGNORECASE
    )

    # ------------------------------------------------
    # FUNCION PARA LIMPIAR NOMBRES
    # ------------------------------------------------

    def limpiar_nombre(nombre):

        nombre = limpiar_descripcion_procedimiento(nombre)

        nombre = re.sub(
            r"\s+Y\d{2,3}\s*$",
            "",
            nombre,
            flags=re.IGNORECASE
        )

        return re.sub(r"\s+", " ", nombre).strip().upper()

    # ------------------------------------------------
    # FUNCION PARA IDENTIFICAR CODIGO Y PROCEDIMIENTO
    # ------------------------------------------------

    patron_fila = re.compile(
        r"^\s*(?:(\d+)\s+)?"
        r"([A-Z]{1,3}\d{5,8}|\d{5,8}(?:-\d+)?)"
        r"(?:\s+(.+?))?\s*$",
        re.IGNORECASE
    )

    # =================================================
    # PARTE A: LEER APARTADO CIRUGIAS
    # =================================================

    inicio_tabla = None

    for i, linea in enumerate(lineas):

        if patron_cirugias.match(linea):

            inicio_tabla = i + 1
            break

    if inicio_tabla is not None:

        for linea in lineas[inicio_tabla:]:

            # La tabla termina cuando comienza la descripcion.
            if patron_descripcion.match(linea):

                break

            coincidencia = patron_fila.match(linea)

            if not coincidencia:
                continue

            cantidad = int(coincidencia.group(1) or 1)

            codigo = coincidencia.group(2).strip().upper()

            nombre = limpiar_nombre(
                coincidencia.group(3) or ""
            )

            if not nombre:
                continue

            procedimientos_tabla.append({
                "codigo": codigo,
                "procedimiento": nombre,
                "cantidad": cantidad
            })

    # =================================================
    # PARTE B: LEER DESCRIPCION CIRUGIA
    # =================================================

    dentro_descripcion = False

    for linea in lineas:

        if patron_descripcion.match(linea):

            dentro_descripcion = True
            continue

        if not dentro_descripcion:
            continue

        linea_normalizada = normalizar_texto(linea)

        # Final de la seccion.
        if re.match(
            r"^\s*dx\s+(?:preoperatorio|postoperatorio)\b",
            linea_normalizada
        ):
            break

        if re.match(
            r"^\s*(?:tipo de herida|cantidad de sangrado|"
            r"realizacion acto quirurgico|complicaciones)\b",
            linea_normalizada
        ):
            break

        coincidencia = patron_fila.match(linea)

        if not coincidencia:
            continue

        cantidad = int(coincidencia.group(1) or 1)

        codigo = coincidencia.group(2).strip().upper()

        nombre = limpiar_nombre(
            coincidencia.group(3) or ""
        )

        if not nombre:
            continue

        procedimientos_descripcion.append({
            "codigo": codigo,
            "procedimiento": nombre,
            "cantidad": cantidad
        })

    # =================================================
    # PARTE C: UNIFICAR AMBAS FUENTES
    # =================================================

    procedimientos = list(procedimientos_tabla)

    codigos_existentes = {
        p["codigo"]
        for p in procedimientos
        if p["codigo"]
    }

    for descripcion in procedimientos_descripcion:

        codigo = descripcion["codigo"]

        # Si el codigo ya existe en la tabla,
        # no volver a agregarlo.

        if codigo in codigos_existentes:
            continue

        procedimientos.append(descripcion)

        codigos_existentes.add(codigo)

    # =================================================
    # RESULTADO FINAL
    # =================================================

    return procedimientos

  


# ==================================================
# MODULO 6: RESUMEN DE CALIDAD DE DATOS
# ==================================================

def generar_resumen_campos(df):

    campos_generales = [
        "ESPECIALISTA",
        "FECHA",
        "IDENTIF",
        "PACIENTE",
        "ENTIDAD",
    ]

    resumen = []

    for campo in campos_generales:

        completos = int(
            df[campo]
            .fillna("")
            .astype(str)
            .str.strip()
            .ne("")
            .sum()
        )

        resumen.append({
            "CAMPO": campo,
            "REGISTROS COMPLETOS": completos,
            "REGISTROS VACIOS": len(df) - completos,
        })

    return pd.DataFrame(resumen)


# ==================================================
# MODULO 7: FORMATO DEL EXCEL
# ==================================================

def aplicar_formato_excel(writer):

    hoja = writer.sheets["Extracto"]

    anchos = {
        "A": 30,
        "B": 15,
        "C": 25,
        "D": 35,
        "E": 18,
        "F": 70,
        "G": 35,
        "H": 18,
        "I": 35,
    }

    for columna, ancho in anchos.items():

        hoja.column_dimensions[columna].width = ancho

    hoja.freeze_panes = "A2"

    hoja.auto_filter.ref = hoja.dimensions

    # Encabezados del extracto.
    for celda in hoja[1]:

        celda.font = Font(
            bold=True,
            color="FFFFFF"
        )

        celda.fill = PatternFill(
            fill_type="solid",
            fgColor="008C68"
        )

        celda.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    hoja.row_dimensions[1].height = 25

    # Formato del resumen.
    hoja_resumen = writer.sheets["Resumen"]

    hoja_resumen.column_dimensions["A"].width = 35
    hoja_resumen.column_dimensions["B"].width = 25
    hoja_resumen.column_dimensions["C"].width = 22

    for celda in hoja_resumen[1]:

        celda.font = Font(
            bold=True,
            color="FFFFFF"
        )

        celda.fill = PatternFill(
            fill_type="solid",
            fgColor="008C68"
        )

    hoja_resumen.freeze_panes = "A2"


# ==================================================
# MODULO 8: CONSOLIDACION DE TODOS LOS DIAS
# ==================================================

def generar_extracto_excel():

    print("\n" + "=" * 55)
    print("       CONSOLIDACION DE NOTAS QUIRURGICAS")
    print("=" * 55)

    todos_los_registros = []

    total_pdfs = 0
    pdf_con_procedimientos = 0
    pdf_sin_procedimientos = 0
    total_errores = 0
    total_procedimientos = 0

    # ----------------------------------------------
    # RECORRER LAS 31 CARPETAS
    # ----------------------------------------------

    for numero_dia in range(1, 32):

        carpeta_dia = CARPETA_PDFS / f"dia{numero_dia}"

        archivos_pdf = obtener_archivos_pdf(carpeta_dia)

        if not archivos_pdf:
            continue

        print(
            f"\nProcesando dia{numero_dia}: "
            f"{len(archivos_pdf)} PDF"
        )

        # ------------------------------------------
        # PROCESAR CADA PDF
        # ------------------------------------------

        for archivo in archivos_pdf:

            total_pdfs += 1

            try:

                texto = extraer_texto_pdf(archivo)

                if not texto.strip():

                    print(
                        f"PDF sin texto extraible: {archivo.name}"
                    )

                    pdf_sin_procedimientos += 1
                    continue

                datos_generales = extraer_datos_generales(texto)

                procedimientos = extraer_procedimientos(texto)

                if not procedimientos:

                    print(
                        f"Sin procedimientos: "
                        f"dia{numero_dia}/{archivo.name}"
                    )

                    pdf_sin_procedimientos += 1
                    continue

                pdf_con_procedimientos += 1

                cantidad_procedimientos_pdf = 0

                # ----------------------------------
                # GENERAR FILAS SEGUN CANTIDAD
                # ----------------------------------

                for procedimiento in procedimientos:

                    cantidad = procedimiento.get(
                        "cantidad",
                        1
                    )

                    for _ in range(cantidad):

                        registro = {
                            **datos_generales,

                            "CODIGO": procedimiento["codigo"],

                            "PROCEDIMIENTOS": procedimiento["procedimiento"],

                            "DIA_CARPETA": f"dia{numero_dia}",

                            "ARCHIVO_PDF": archivo.relative_to(
                                carpeta_dia
                            ).as_posix()
                        }

                        todos_los_registros.append(registro)

                        total_procedimientos += 1
                        cantidad_procedimientos_pdf += 1

                # ----------------------------------
                # REVISAR CAMPOS GENERALES VACIOS
                # ----------------------------------

                campos_vacios = [
                    campo
                    for campo, valor in datos_generales.items()
                    if not str(valor).strip()
                ]

                if campos_vacios:

                    print(
                        f"{archivo.name}: "
                        f"{cantidad_procedimientos_pdf} registros. "
                        f"Revisar campos: "
                        f"{', '.join(campos_vacios)}"
                    )

                else:

                    print(
                        f"{archivo.name}: "
                        f"{cantidad_procedimientos_pdf} registros exportables."
                    )

            except Exception as error:

                total_errores += 1

                print(
                    f"Error procesando {archivo.name}: "
                    f"{type(error).__name__}: {error}"
                )

    # ----------------------------------------------
    # VALIDAR RESULTADOS
    # ----------------------------------------------

    if not todos_los_registros:

        print("\nNo se encontraron procedimientos para exportar.")

        print("PDF revisados:", total_pdfs)
        print("PDF sin procedimientos:", pdf_sin_procedimientos)
        print("Errores:", total_errores)

        return

    # ----------------------------------------------
    # CREAR DATAFRAME CONSOLIDADO
    # ----------------------------------------------

    df = pd.DataFrame(
        todos_los_registros,
        columns=COLUMNAS_EXTRACTO
    )

    # ----------------------------------------------
    # ASEGURAR IDENTIFICACIONES COMO TEXTO
    # ----------------------------------------------

    df["IDENTIF"] = df["IDENTIF"].fillna("").astype(str)

    # ----------------------------------------------
    # CREAR RESUMEN DE CALIDAD
    # ----------------------------------------------

    resumen_campos = generar_resumen_campos(df)

    # ----------------------------------------------
    # CREAR RESUMEN GENERAL
    # ----------------------------------------------

    resumen_general = pd.DataFrame({

        "INDICADOR": [
            "Mes procesado",
            "PDF encontrados",
            "PDF con procedimientos",
            "PDF sin procedimientos",
            "Errores de procesamiento",
            "Procedimientos exportados",
        ],

        "CANTIDAD": [
            MES_PROCESO,
            total_pdfs,
            pdf_con_procedimientos,
            pdf_sin_procedimientos,
            total_errores,
            total_procedimientos,
        ]

    })

    # ----------------------------------------------
    # UNIR LOS RESUMENES
    # ----------------------------------------------

    resumen_final = pd.concat(
        [
            resumen_general,
            pd.DataFrame([{}]),
            resumen_campos
        ],
        ignore_index=True
    )

    # ----------------------------------------------
    # CREAR CARPETA DE RESULTADOS
    # ----------------------------------------------

    CARPETA_RESULTADOS_MES.mkdir(
        parents=True,
        exist_ok=True
    )

    ruta_salida = CARPETA_RESULTADOS_MES / NOMBRE_EXCEL

    # ----------------------------------------------
    # EXPORTAR UN UNICO EXCEL
    # ----------------------------------------------

    try:

        with pd.ExcelWriter(
            ruta_salida,
            engine="openpyxl"
        ) as writer:

            df.to_excel(
                writer,
                sheet_name="Extracto",
                index=False
            )

            resumen_final.to_excel(
                writer,
                sheet_name="Resumen",
                index=False
            )

            aplicar_formato_excel(writer)

        print("\n" + "=" * 55)
        print("           PROCESO FINALIZADO")
        print("=" * 55)

        print("PDF encontrados:", total_pdfs)
        print("PDF con procedimientos:", pdf_con_procedimientos)
        print("PDF sin procedimientos:", pdf_sin_procedimientos)
        print("Errores:", total_errores)
        print("Procedimientos exportados:", total_procedimientos)

        print("\nUN SOLO ARCHIVO EXCEL GENERADO:")
        print(ruta_salida)

    except PermissionError:

        print(
            "\nNo se pudo guardar el Excel. "
            "Cierra el archivo si esta abierto en Excel "
            "y vuelve a ejecutar."
        )

    except Exception as error:

        print(
            "\nError al generar el Excel:",
            type(error).__name__,
            error
        )


# ==================================================
# EJECUCION PRINCIPAL
# ==================================================

if __name__ == "__main__":

    if verificar_estructura():

        generar_extracto_excel()
