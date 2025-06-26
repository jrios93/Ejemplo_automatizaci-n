import pdfplumber
import re
import pyautogui
import pyperclip
import time
import os


def extraer_datos(path_pdf):
    with pdfplumber.open(path_pdf) as pdf:
        texto = "\n".join([p.extract_text() for p in pdf.pages])

    print("=== TEXTO EXTRAÍDO ===")
    print(texto[:500] + "...")  # Mostrar primeros 500 caracteres para debug
    print("\n" + "="*50 + "\n")

    # Extraer número de serie completo (F001-0000033)
    referencia_match = re.search(r'REFERENCIA\s*:\s*(F\d{3}-\d{7,8})', texto)
    referencia = referencia_match.group(
        1) if referencia_match else "No encontrada"

    serie_numero_match = re.search(r'(F\d{3}-\d{7})', texto)
    if serie_numero_match:
        serie_completa = serie_numero_match.group(1)
        serie = serie_completa.split('-')[0]  # F001
        numero = serie_completa.split('-')[1]  # 0000033
    else:
        # Buscar por separado si no encuentra el patrón completo
        serie_match = re.search(r'(F\d{3})', texto)
        numero_match = re.search(r'F\d{3}-(\d{7})', texto)
        serie = serie_match.group(1) if serie_match else 'F001'
        numero = numero_match.group(1) if numero_match else '0000000'

    # Extraer fecha (dd/mm/yyyy)
    fecha_match = re.search(r'(\d{2}/\d{2}/\d{4})', texto)
    if fecha_match:
        fecha_original = fecha_match.group(1)
        # Convertir a formato ddmmyyyy para Contasis
        fecha_partes = fecha_original.split('/')
        fecha_contasis = fecha_partes[0] + fecha_partes[1] + fecha_partes[2]
    else:
        fecha_contasis = '23062025'  # Fecha por defecto
        fecha_original = '23/06/2025'

    # RUC fijo del cliente
    ruc = '20568033354'

    # Extraer productos de forma más robusta
    productos = []
    # Buscar la tabla de productos
    lineas = texto.split('\n')
    en_tabla_productos = False

    for i, linea in enumerate(lineas):
        # Detectar inicio de tabla de productos
        if 'CANT.' in linea and 'CODIGO' in linea and 'DESCRIPCION' in linea and 'P. UNIT.' in linea:
            en_tabla_productos = True
            continue

        # Si estamos en la tabla de productos
        if en_tabla_productos:
            # Detectar fin de tabla
            if 'SUB TOTAL' in linea or 'SON:' in linea or 'TOTAL' in linea:
                break

            # Buscar líneas de productos
            linea = linea.strip()
            if linea and not linea.startswith(('CONSIDERACIONES', 'GRACIAS', 'Autorizado')):
                # Buscar código de 12 dígitos en la línea
                codigo_match = re.search(r'\b(\d{12})\b', linea)
                if codigo_match:
                    codigo = codigo_match.group(1)
                    # Extraer cantidad (al inicio de la línea)
                    cantidad_match = re.match(r'^(\d+\.?\d*)', linea)
                    cantidad = cantidad_match.group(
                        1) if cantidad_match else '1.00'

                    # Extraer precios (buscar números con formato de precio)
                    precios = re.findall(
                        r'\b(\d{1,3}(?:,\d{3})*\.?\d{2})\b', linea)
                    precio_unitario = precios[-2] if len(precios) >= 2 else (
                        precios[0] if precios else '0.00')
                    precio_unitario = precio_unitario.replace(',', '')

                    # Extraer descripción
                    descripcion = linea
                    descripcion = re.sub(r'^\d+\.?\d*\s+', '', descripcion)
                    descripcion = descripcion.replace(codigo, '').strip()
                    for precio in precios:
                        descripcion = descripcion.replace(precio, '').strip()
                    descripcion = re.sub(r'\s+', ' ', descripcion).strip()

                    productos.append({
                        "cantidad": cantidad,
                        "codigo": codigo,
                        "descripcion": descripcion,
                        "precio_unitario": precio_unitario
                    })

    # Si no se encontraron productos con el método anterior, usar método alternativo
    if not productos:
        for linea in lineas:
            if re.search(r'\d{12}', linea):
                codigo_match = re.search(r'(\d{12})', linea)
                if codigo_match:
                    codigo = codigo_match.group(1)
                    numeros = re.findall(r'\d+[.,]\d{2}', linea)
                    productos.append({
                        "cantidad": "1.00",
                        "codigo": codigo,
                        "descripcion": "PRODUCTO - Verificar descripción manualmente",
                        "precio_unitario": numeros[-1].replace(',', '.') if numeros else '0.00',
                        "linea_original": linea.strip()
                    })

    return {
        "serie": serie,
        "numero": numero,
        "fecha": fecha_contasis,
        "fecha_original": fecha_original,
        "ruc": ruc,
        "referencia": referencia,  # Ahora devuelve el string, no el objeto regex
        "productos": productos
    }


def escribir_en_contasis(datos):
    print("Iniciando en 5 segundos... Coloca el foco en Contasis!")
    time.sleep(5)
    pyautogui.PAUSE = 0.2

    # DEPURACIÓN: Mostrar lo que se enviará
    print("Datos a ingresar:")
    print(f"Serie: {datos['serie']}")
    print(f"Número: {datos['numero']}")
    print(f"Fecha: {datos['fecha']}")
    print(f"RUC: {datos['ruc']}")
    print(f"Productos: {len(datos['productos'])}")

    try:
        # Ingresar serie y número
        pyautogui.write(datos["serie"])
        pyautogui.press("tab")
        pyautogui.write(datos["numero"])
        pyautogui.press("tab")

        # Ingresar RUC y datos adicionales
        pyautogui.write(datos["ruc"])
        pyautogui.press("tab", presses=2)
        pyautogui.write("000")
        pyautogui.press("tab")

        # Ingresar fecha
        pyautogui.write(datos["fecha"])
        pyautogui.write("CRE")
        pyautogui.press("tab", presses=8)
        pyautogui.press("enter")
        time.sleep(1)

        # Ingresar productos
        for i, producto in enumerate(datos["productos"]):
            pyautogui.write(producto["codigo"])
            pyautogui.press("enter")
            time.sleep(0.5)
            pyautogui.write(producto["cantidad"])
            pyautogui.press("tab")
            time.sleep(0.5)
            pyautogui.press("tab")
            pyautogui.write(producto["precio_unitario"])
            if i < len(datos["productos"]) - 1:
                pyautogui.hotkey("ctrl", "v")
                time.sleep(1)

        print("✅ Proceso completado automáticamente.")
    except Exception as e:
        print(f"❌ Error durante la automatización: {e}")
        print("Verifica la posición del cursor en Contasis.")


def procesar_archivo_unico():
    """Procesa un solo archivo PDF"""
    archivo_pdf = input("Ingresa el nombre del PDF (ej: F001-0033): ").strip()
    nombre_completo = f'NC {archivo_pdf}.pdf'

    try:
        datos = extraer_datos(nombre_completo)
        mostrar_datos_extraidos(datos)

        respuesta = input(
            "\n¿Deseas continuar con la automatización en Contasis? (s/n): ")
        if respuesta.lower() == 's':
            escribir_en_contasis(datos)
            print("Proceso completado. Verifica los resultados en Contasis.")
            print(f"Referencia: {datos['referencia']}")
        else:
            print("Extracción completada. Datos listos para uso manual.")

    except FileNotFoundError:
        print(f"❌ Error: No se encontró el archivo '{nombre_completo}'")
        print("Asegúrate de que el archivo esté en la misma carpeta que este script.")
    except Exception as e:
        print(f"❌ Error al procesar el PDF: {str(e)}")


def procesar_multiples_archivos():
    """Procesa múltiples archivos PDF de forma consecutiva"""
    archivos_procesados = 0

    while True:
        print(f"\n{'='*50}")
        print(f"🔄 PROCESANDO ARCHIVO #{archivos_procesados + 1}")
        print(f"{'='*50}")

        archivo_pdf = input(
            "Ingresa el nombre del PDF (ej: F001-0033) o 'salir' para terminar: ").strip()

        if archivo_pdf.lower() in ['salir', 'exit', 'n', 'no']:
            break

        nombre_completo = f'NC {archivo_pdf}.pdf'

        try:
            datos = extraer_datos(nombre_completo)
            mostrar_datos_extraidos(datos)

            respuesta = input(
                "\n¿Deseas continuar con la automatización en Contasis para este archivo? (s/n): ")
            if respuesta.lower() == 's':
                escribir_en_contasis(datos)
                print("✅ Archivo procesado correctamente en Contasis.")
                print(f"Referencia: {datos['referencia']}")
            else:
                print("ℹ️  Datos extraídos, pero no se ejecutó la automatización.")

            archivos_procesados += 1

            # Preguntar si quiere continuar con otro archivo
            continuar = input("\n¿Deseas procesar otro archivo? (s/n): ")
            if continuar.lower() != 's':
                break

        except FileNotFoundError:
            print(f"❌ Error: No se encontró el archivo '{nombre_completo}'")
            print("Asegúrate de que el archivo esté en la misma carpeta.")
            continue
        except Exception as e:
            print(f"❌ Error al procesar el PDF: {str(e)}")
            continue

    print(
        f"\n🎉 Proceso finalizado. Se procesaron {archivos_procesados} archivo(s).")


def mostrar_datos_extraidos(datos):
    """Muestra los datos extraídos de forma organizada"""
    print("\n=== DATOS EXTRAÍDOS ===")
    print(f"Serie completa: {datos['serie']}-{datos['numero']}")
    print(f"Serie: {datos['serie']}")
    print(f"Número: {datos['numero']}")
    print(f"Fecha original: {datos['fecha_original']}")
    print(f"Fecha (formato Contasis): {datos['fecha']}")
    print(f"RUC: {datos['ruc']}")
    print(f"Referencia: {datos['referencia']}")  # Ahora muestra correctamente
    print(f"Productos encontrados: {len(datos['productos'])}")

    for i, p in enumerate(datos["productos"]):
        print(f" {i+1}. Código: {p['codigo']}")
        print(f"    Descripción: {p['descripcion']}")
        print(f"    Cantidad: {p['cantidad']}")
        print(f"    Precio Unitario: S/ {p['precio_unitario']}")
        if 'linea_original' in p:
            print(f"    Línea original: {p['linea_original']}")

    if not datos['productos']:
        print("⚠️  No se detectaron productos automáticamente.")
        print("    Revisa el PDF y verifica el formato de la tabla de productos.")


# EJECUCIÓN PRINCIPAL MEJORADA
if __name__ == "__main__":
    print("🚀 EXTRACTOR DE NOTAS DE CRÉDITO")
    print("Desarrollado para procesamiento automático en Contasis")
    print("="*60)

    modo = input(
        "\n¿Qué deseas hacer?\n"
        "1. Procesar UN archivo\n"
        "2. Procesar MÚLTIPLES archivos\n"
        "Selecciona (1 o 2): "
    ).strip()

    if modo == "1":
        procesar_archivo_unico()
    elif modo == "2":
        procesar_multiples_archivos()
    else:
        print("❌ Opción no válida. Ejecuta el programa nuevamente.")
