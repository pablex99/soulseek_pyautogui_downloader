import pytesseract
import cv2
from difflib import SequenceMatcher
import pandas as pd
import pyautogui
import time
import pygetwindow as gw
import os
import re

# 1. Leer el Excel con canciones y autores
def leer_excel(ruta_excel):
    df = pd.read_excel(ruta_excel)
    canciones = []
    for idx, row in df.iterrows():
        cancion = str(row.iloc[0]).strip()
        autor = str(row.iloc[1]).strip()
        # Si la celda de canción o autor está vacía, termina la carga
        if not cancion or not autor or cancion.lower() == 'nan' or autor.lower() == 'nan':
            print(f"Fin de la lista detectado en la fila {idx+1}.")
            break
        canciones.append((cancion, autor))
    return canciones

# 2. Poner SoulseekQt al frente
def enfocar_soulseek():
    for w in gw.getAllTitles():
        if 'soulseek' in w.lower():
            win = gw.getWindowsWithTitle(w)[0]
            try:
                win.activate()
            except Exception as e:
                print(f"[ADVERTENCIA] No se pudo activar la ventana: {e}")
            try:
                win.maximize()
            except Exception as e:
                print(f"[ADVERTENCIA] No se pudo maximizar la ventana: {e}")
            time.sleep(1)
            return True
    return False

# 3. Pedir al usuario que seleccione las zonas clave
def pedir_posiciones():
    print("Coloca el mouse sobre la barra de búsqueda. Tienes 5 segundos...")
    time.sleep(5)
    barra_busqueda = pyautogui.position()
    print(f"Posición barra de búsqueda: {barra_busqueda}")
    print("Coloca el mouse sobre la pestaña/sección de descargas. Tienes 5 segundos...")
    time.sleep(5)
    seccion_descargas = pyautogui.position()
    print(f"Posición sección descargas: {seccion_descargas}")
    print("Coloca el mouse sobre la pestaña/sección de buscar. Tienes 5 segundos...")
    time.sleep(5)
    seccion_buscar = pyautogui.position()
    print(f"Posición sección buscar: {seccion_buscar}")
    print("Coloca el mouse sobre la PRIMERA OPCIÓN para descargar tras una búsqueda. Tienes 5 segundos...")
    time.sleep(5)
    primera_opcion = pyautogui.position()
    print(f"Posición primera opción de descarga: {primera_opcion}")
    return barra_busqueda, seccion_descargas, seccion_buscar, primera_opcion

# 4. Automatizar búsqueda y descarga
def automatizar(canciones, barra_busqueda, seccion_descargas, seccion_buscar, primera_opcion):
    import numpy as np
    import re

    def detectar_alto_opcion(img_res):
        # Convierte a escala de grises
        gray = cv2.cvtColor(img_res, cv2.COLOR_RGB2GRAY)
        # Aplica filtro de bordes
        edges = cv2.Canny(gray, 50, 150, apertureSize=3)
        # Suma de bordes por fila
        sumas = np.sum(edges, axis=1)
        # Umbral para detectar líneas horizontales (divisores de filas)
        umbral = np.max(sumas) * 0.5
        indices = np.where(sumas > umbral)[0]
        # Filtra líneas muy cercanas (solo cuenta saltos grandes)
        if len(indices) < 2:
            return 32  # fallback
        saltos = np.diff(indices)
        saltos = saltos[saltos > 5]  # ignora líneas muy pegadas
        if len(saltos) == 0:
            return 32  # fallback
        alto_prom = int(np.median(saltos))
        print(f"[DEBUG] Alto de opción detectado automáticamente: {alto_prom} px")
        return alto_prom

    def normalizar(texto):
        # Quitar símbolos, guiones, guiones bajos, números, etc. y pasar a minúsculas
        return re.sub(r'[^a-záéíóúüñ ]', '', texto.lower())

    descargadas = []  # lista de tuplas (busqueda, texto_detectado)
    no_descargadas = []  # lista de strings (busqueda)
    # Carpeta de descargas personalizada
    carpeta_descargas = r'C:\Users\terra\OneDrive\Escritorio\Descargas Automatizadas\Completadas\complete'

    def archivos_en_descargas():
        archivos = set()
        for root, dirs, files in os.walk(carpeta_descargas):
            for f in files:
                archivos.add(os.path.join(root, f))
        return archivos

    # --- NUEVO: Mensaje para dejar solo Soulseek en pantalla ---
    print("\n[ATENCIÓN] Asegúrate de que SOLO la ventana de Soulseek esté visible y maximizada en pantalla completa. Tienes 5 segundos...")
    time.sleep(5)
    # Maximizar Soulseek solo antes de la primera descarga
    enfocar_soulseek()

    for idx, (cancion, autor) in enumerate(canciones):
            # ...existing code...
        print(f"Buscando y descargando: {cancion} - {autor}")
        # Ir a la sección buscar
        pyautogui.click(seccion_buscar)
        time.sleep(0.5)
        # Click en barra de búsqueda
        pyautogui.click(barra_busqueda)
        time.sleep(0.2)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('backspace')
        # Agregar filtros para mp3 y 320 kbps
        busqueda = f"{cancion} {autor} mp3 320"
        pyautogui.write(busqueda, interval=0.05)
        pyautogui.press('enter')
        time.sleep(9)  # Esperar más tiempo para que carguen todos los resultados

        # Tomar screenshot de una franja vertical ajustada: solo ampliar el borde derecho (ancho)
        ancho = 1200  # Mucho más ancho para que el límite derecho esté bastante más hacia la derecha
        desplazamiento_izq = 31  # Mover el borde izquierdo más a la izquierda
        x0 = primera_opcion.x - desplazamiento_izq
        y0 = primera_opcion.y - 10  # Un poco más arriba de la primera opción
        alto = 600   # Valor original que funcionaba correctamente
        screenshot_opciones = pyautogui.screenshot(region=(x0, y0, ancho, alto))

        # OCR para extraer texto solo de la franja vertical de opciones
        import pytesseract
        import cv2
        import numpy as np
        def obtener_opciones_y_cajas(screenshot_img):
            img = cv2.cvtColor(np.array(screenshot_img), cv2.COLOR_RGB2BGR)
            texto_opciones = pytesseract.image_to_string(img)
            lineas_opciones = [linea.strip() for linea in texto_opciones.splitlines() if linea.strip()]
            ocr_data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
            cajas_lineas = []
            for i in range(len(ocr_data['level'])):
                if ocr_data['text'][i].strip():
                    cajas_lineas.append({
                        'text': ocr_data['text'][i].strip(),
                        'left': ocr_data['left'][i],
                        'top': ocr_data['top'][i],
                        'width': ocr_data['width'][i],
                        'height': ocr_data['height'][i]
                    })
            return lineas_opciones, cajas_lineas, img

        lineas_opciones, cajas_lineas, img = obtener_opciones_y_cajas(screenshot_opciones)

        # Guardar log de opciones y la opción más parecida a la búsqueda (sin '320')
        with open('soulseek_opciones_log.txt', 'a', encoding='utf-8') as flog:
            flog.write(f"\n--- Opciones para búsqueda: {busqueda} ---\n")
            if lineas_opciones:
                # Herramienta de debugging: imagen con marcas
                debug_img = img.copy()
                debug_img_all = img.copy()
                debug_clic_info = None
                for caja in cajas_lineas:
                    pt1 = (caja['left'], caja['top'])
                    pt2 = (caja['left'] + caja['width'], caja['top'] + caja['height'])
                    cv2.rectangle(debug_img_all, pt1, pt2, (0,0,255), 2)
                def limpiar_opcion(op):
                    op = op.strip()
                    if op.lower().endswith('.mp3'):
                        op = op[:-4].strip()
                    op = re.sub(r'[.,;]', '', op)
                    return op
                opciones_limpias = [limpiar_opcion(opcion) for opcion in lineas_opciones]
                for op in opciones_limpias:
                    flog.write(op + '\n')
                def normalizar_texto(texto):
                    return re.sub(r'[^a-záéíóúüñ ]', '', texto.lower())
                busqueda_sin_320_mp3 = busqueda.replace('320', '').replace('mp3', '').strip()
                palabras_busqueda = [w for w in normalizar_texto(busqueda_sin_320_mp3).split() if w != '']
                palabras_busqueda_set = set(palabras_busqueda)
                # Refinado: máxima prioridad a coincidencia exacta de palabras (sin mp3/320),
                # luego a las que tengan todas las palabras y menos extras, y por último score aproximado
                mejor_opcion = None
                mejor_opcion_extras = None
                min_extras = None
                for op in opciones_limpias:
                    palabras_opcion = [w for w in normalizar_texto(op).split() if w != '']
                    if set(palabras_opcion) == palabras_busqueda_set and len(palabras_opcion) == len(palabras_busqueda):
                        mejor_opcion = op
                        break
                if not mejor_opcion:
                    for op in opciones_limpias:
                        palabras_opcion = [w for w in normalizar_texto(op).split() if w != '']
                        if palabras_busqueda_set.issubset(set(palabras_opcion)):
                            extras = len(palabras_opcion) - len(palabras_busqueda)
                            if min_extras is None or extras < min_extras:
                                mejor_opcion_extras = op
                                min_extras = extras
                    if mejor_opcion_extras:
                        mejor_opcion = mejor_opcion_extras
                if not mejor_opcion:
                    def score_opcion(op):
                        palabras_opcion = [w for w in normalizar_texto(op).split() if w != '']
                        matches = sum(1 for palabra in palabras_busqueda if palabra in palabras_opcion)
                        extras = max(0, len(palabras_opcion) - len(palabras_busqueda))
                        score = matches * 3 - extras
                        return score
                    mejor_opcion = max(opciones_limpias, key=score_opcion)
                flog.write(f"[Mejor coincidencia refinada]: {mejor_opcion}\n")
            else:
                flog.write("[ADVERTENCIA] No se detectaron opciones por OCR. Revisa la imagen soulseek_opciones_{idx+1}.png para depurar la región capturada.\n")

        # Descargar la mejor opción encontrada por el algoritmo
        if lineas_opciones:
            # --- NUEVO: Antes de hacer clic, volver a tomar screenshot y OCR, y verificar si la mejor opción sigue presente ---
            try:
                pyautogui.moveTo(barra_busqueda)  # Mueve el mouse fuera de la zona de opciones para evitar hover
                time.sleep(0.5)
            except pyautogui.FailSafeException:
                print("[ADVERTENCIA] PyAutoGUI fail-safe activado: el mouse se movió a una esquina de la pantalla. Se omite esta canción y se continúa con la siguiente.")
                continue
            screenshot_opciones2 = pyautogui.screenshot(region=(x0, y0, ancho, alto))
            lineas_opciones2, cajas_lineas2, img2 = obtener_opciones_y_cajas(screenshot_opciones2)
            opciones_limpias2 = [re.sub(r'[.,;]', '', op.strip()[:-4].strip()) if op.strip().lower().endswith('.mp3') else re.sub(r'[.,;]', '', op.strip()) for op in lineas_opciones2]
            def normalizar_texto(texto):
                return re.sub(r'[^a-záéíóúüñ ]', '', texto.lower())
            palabras_busqueda = [w for w in normalizar_texto(busqueda.replace('320', '').replace('mp3', '').strip()).split() if w != '']
            palabras_busqueda_set = set(palabras_busqueda)
            # ¿Sigue la mejor opción?
            mejor_opcion_presente = False
            for op in opciones_limpias2:
                palabras_opcion = [w for w in normalizar_texto(op).split() if w != '']
                if set(palabras_opcion) == palabras_busqueda_set and len(palabras_opcion) == len(palabras_busqueda):
                    if op == mejor_opcion:
                        mejor_opcion_presente = True
                        break
            if not mejor_opcion_presente:
                # Recalcular mejor opción con las nuevas opciones y registrar en el log
                with open('soulseek_opciones_log.txt', 'a', encoding='utf-8') as flog:
                    flog.write(f"\n[Repetición de proceso: nuevas opciones detectadas tras refresco OCR]\n")
                    for op in opciones_limpias2:
                        flog.write(op + '\n')
                mejor_opcion = None
                mejor_opcion_extras = None
                min_extras = None
                for op in opciones_limpias2:
                    palabras_opcion = [w for w in normalizar_texto(op).split() if w != '']
                    if set(palabras_opcion) == palabras_busqueda_set and len(palabras_opcion) == len(palabras_busqueda):
                        mejor_opcion = op
                        break
                if not mejor_opcion:
                    for op in opciones_limpias2:
                        palabras_opcion = [w for w in normalizar_texto(op).split() if w != '']
                        if palabras_busqueda_set.issubset(set(palabras_opcion)):
                            extras = len(palabras_opcion) - len(palabras_busqueda)
                            if min_extras is None or extras < min_extras:
                                mejor_opcion_extras = op
                                min_extras = extras
                    if mejor_opcion_extras:
                        mejor_opcion = mejor_opcion_extras
                if not mejor_opcion:
                    def score_opcion(op):
                        palabras_opcion = [w for w in normalizar_texto(op).split() if w != '']
                        matches = sum(1 for palabra in palabras_busqueda if palabra in palabras_opcion)
                        extras = max(0, len(palabras_opcion) - len(palabras_busqueda))
                        score = matches * 3 - extras
                        return score
                    if opciones_limpias2:
                        mejor_opcion = max(opciones_limpias2, key=score_opcion)
                with open('soulseek_opciones_log.txt', 'a', encoding='utf-8') as flog:
                    flog.write(f"[Nueva mejor opción tras repetición]: {mejor_opcion}\n")
                # Actualizar cajas_lineas y debug_img para el nuevo set
                cajas_lineas = cajas_lineas2
                debug_img = img2.copy()
                debug_img_all = img2.copy()
            # Usar el mismo algoritmo de coincidencia para encontrar la mejor opción
            # import re ya está al inicio del script
            def limpiar_opcion(op):
                op = op.strip()
                if op.lower().endswith('.mp3'):
                    op = op[:-4].strip()
                op = re.sub(r'[.,;]', '', op)
                return op
            opciones_limpias = [limpiar_opcion(opcion) for opcion in lineas_opciones]
            def normalizar_texto(texto):
                return re.sub(r'[^a-záéíóúüñ ]', '', texto.lower())
            busqueda_sin_320_mp3 = busqueda.replace('320', '').replace('mp3', '').strip()
            palabras_busqueda = [w for w in normalizar_texto(busqueda_sin_320_mp3).split() if w != '']
            palabras_busqueda_set = set(palabras_busqueda)
            mejor_opcion = None
            for op in opciones_limpias:
                palabras_opcion = [w for w in normalizar_texto(op).split() if w != '']
                if set(palabras_opcion) == palabras_busqueda_set:
                    mejor_opcion = op
                    break
            if not mejor_opcion:
                for op in opciones_limpias:
                    palabras_opcion = [w for w in normalizar_texto(op).split() if w != '']
                    if palabras_busqueda_set.issubset(set(palabras_opcion)):
                        mejor_opcion = op
                        break
            if not mejor_opcion:
                def score_opcion(op):
                    palabras_opcion = [w for w in normalizar_texto(op).split() if w != '']
                    matches = sum(1 for palabra in palabras_busqueda if palabra in palabras_opcion)
                    extras = max(0, len(palabras_opcion) - len(palabras_busqueda))
                    score = matches * 3 - extras
                    return score
                mejor_opcion = max(opciones_limpias, key=score_opcion)
            # Mejorar: filtrar solo cajas en la zona de canciones (donde hay varias alineadas verticalmente)
            # 1. Agrupar cajas por eje X (columna) para encontrar la columna con más cajas (zona de canciones)
            if cajas_lineas:
                from collections import Counter
                # Considerar el centro X de cada caja
                centros_x = [caja['left'] + caja['width']//2 for caja in cajas_lineas]
                # Agrupar por bins de 40px para tolerancia de alineación
                bins = [((x//40)*40) for x in centros_x]
                bin_counts = Counter(bins)
                if bin_counts:
                    columna_canciones = bin_counts.most_common(1)[0][0]
                    # Filtrar solo cajas en esa columna (±25px de tolerancia)
                    cajas_canciones = [caja for caja in cajas_lineas if abs((caja['left'] + caja['width']//2) - columna_canciones) <= 25]
                else:
                    cajas_canciones = cajas_lineas
            else:
                cajas_canciones = []

            # Buscar todas las cajas OCR que forman parte del nombre de la mejor opción (similitud alta)
            from difflib import SequenceMatcher
            cajas_mejor_opcion = []
            for caja in cajas_canciones:
                texto_caja = limpiar_opcion(caja['text'])
                score = SequenceMatcher(None, texto_caja, mejor_opcion).ratio()
                if score >= 0.6:
                    cajas_mejor_opcion.append(caja)

            if cajas_mejor_opcion:
                # Calcular el centro vertical de cada caja
                y_centros = [caja['top'] + caja['height']//2 for caja in cajas_mejor_opcion]
                # Calcular promedio y desviación estándar
                y_prom = np.mean(y_centros)
                y_std = np.std(y_centros)
                # Filtrar outliers: solo considerar cajas cuyo centro esté a menos de 1.2 desviaciones estándar del promedio
                cajas_filtradas = [caja for caja, y in zip(cajas_mejor_opcion, y_centros) if abs(y - y_prom) <= max(8, 1.2*y_std)]
                if cajas_filtradas:
                    y_centros_filtrados = [caja['top'] + caja['height']//2 for caja in cajas_filtradas]
                    x_centros_filtrados = [caja['left'] + caja['width']//2 for caja in cajas_filtradas]
                    y_click_rel = int(np.mean(y_centros_filtrados))
                    x_click_rel = int(np.mean(x_centros_filtrados))
                    # Dibuja rectángulo verde sobre el bounding box filtrado
                    min_top = min(caja['top'] for caja in cajas_filtradas)
                    max_bottom = max(caja['top'] + caja['height'] for caja in cajas_filtradas)
                    min_left = min(caja['left'] for caja in cajas_filtradas)
                    max_right = max(caja['left'] + caja['width'] for caja in cajas_filtradas)
                    cv2.rectangle(debug_img, (min_left, min_top), (max_right, max_bottom), (0,255,0), 2)
                    cv2.rectangle(debug_img_all, (min_left, min_top), (max_right, max_bottom), (0,255,0), 2)
                    # Dibuja círculo azul en el punto de clic
                    cv2.circle(debug_img, (x_click_rel, y_click_rel), 8, (255,0,0), -1)
                    cv2.circle(debug_img_all, (x_click_rel, y_click_rel), 8, (255,0,0), -1)
                    debug_clic_info = f"[DEBUG] Clic en bounding box mejor opción (sin outliers): ({x_click_rel},{y_click_rel}) abs=({x0 + x_click_rel},{y0 + y_click_rel}) texto='{mejor_opcion}'"
                    cv2.imwrite(f"debug_opcion_{idx+1}.png", debug_img)
                    cv2.imwrite(f"debug_opcion_all_{idx+1}.png", debug_img_all)
                    with open('soulseek_opciones_log.txt', 'a', encoding='utf-8') as flog:
                        flog.write(debug_clic_info + '\n')
                    x_click = x0 + x_click_rel
                    y_click = y0 + y_click_rel
                    print(f"Descargando la mejor opción encontrada: {mejor_opcion}")
                    pyautogui.doubleClick(x_click, y_click)
                    time.sleep(1)
                else:
                    # Si todas las cajas fueron descartadas como outliers, usar el promedio de todas
                    y_click_rel = int(y_prom)
                    x_click_rel = int(np.mean([caja['left'] + caja['width']//2 for caja in cajas_mejor_opcion]))
                    cv2.circle(debug_img, (x_click_rel, y_click_rel), 8, (0,0,255), -1)
                    cv2.circle(debug_img_all, (x_click_rel, y_click_rel), 8, (0,0,255), -1)
                    debug_clic_info = f"[DEBUG] Clic en promedio (todas outliers): ({x_click_rel},{y_click_rel}) abs=({x0 + x_click_rel},{y0 + y_click_rel}) texto='{mejor_opcion}'"
                    cv2.imwrite(f"debug_opcion_{idx+1}.png", debug_img)
                    cv2.imwrite(f"debug_opcion_all_{idx+1}.png", debug_img_all)
                    with open('soulseek_opciones_log.txt', 'a', encoding='utf-8') as flog:
                        flog.write(debug_clic_info + '\n')
                    x_click = x0 + x_click_rel
                    y_click = y0 + y_click_rel
                    print(f"Descargando la mejor opción encontrada: {mejor_opcion}")
                    pyautogui.doubleClick(x_click, y_click)
                    time.sleep(1)
            else:
                # Fallback: buscar la mejor opción entre las líneas limpias nuevamente SOLO si no hay ninguna caja con score aceptable
                idx_mejor = None
                for i, op in enumerate(opciones_limpias):
                    if op == mejor_opcion:
                        idx_mejor = i
                        break
                if idx_mejor is not None:
                    alto_opcion = 32
                    y_mejor_rel = idx_mejor * alto_opcion
                    x_centro_rel = ancho // 2
                    # Dibuja rectángulo rojo estimado sobre la opción en ambas imágenes
                    pt1 = (0, y_mejor_rel)
                    pt2 = (ancho, y_mejor_rel + alto_opcion)
                    cv2.rectangle(debug_img, pt1, pt2, (0,0,255), 2)  # rojo
                    cv2.rectangle(debug_img_all, pt1, pt2, (0,0,255), 2)  # rojo
                    # Dibuja círculo azul en el punto de clic estimado en ambas imágenes
                    cv2.circle(debug_img, (x_centro_rel, y_mejor_rel + alto_opcion // 2), 8, (255,0,0), -1)  # azul
                    cv2.circle(debug_img_all, (x_centro_rel, y_mejor_rel + alto_opcion // 2), 8, (255,0,0), -1)  # azul
                    cv2.imwrite(f"debug_opcion_{idx+1}.png", debug_img)
                    cv2.imwrite(f"debug_opcion_all_{idx+1}.png", debug_img_all)
                    debug_clic_info = f"[DEBUG] Clic fallback: ({x_centro_rel},{y_mejor_rel + alto_opcion // 2}) abs=({x0 + x_centro_rel},{y0 + y_mejor_rel + alto_opcion // 2}) texto='{mejor_opcion}'"
                    with open('soulseek_opciones_log.txt', 'a', encoding='utf-8') as flog:
                        flog.write(debug_clic_info + ' (fallback)\n')
                    y_mejor = y0 + idx_mejor * alto_opcion + alto_opcion // 2
                    x_centro = x0 + ancho // 2
                    print(f"Descargando la mejor opción encontrada (fallback): {mejor_opcion}")
                    pyautogui.doubleClick(x_centro, y_mejor)
                    time.sleep(1)
        else:
            print("No se detectaron opciones para descargar en esta búsqueda.")

    # Al terminar todas las búsquedas, tomar screenshot de la sección de descargas (pantalla completa), aplicar OCR y marcar coincidencias
    print("\nAnalizando sección de descargas con OCR...")
    pyautogui.click(seccion_descargas)
    time.sleep(2)

    screenshot = pyautogui.screenshot()  # Pantalla completa
    screenshot.save("descargas_soulseek.png")
    img = cv2.cvtColor(cv2.imread("descargas_soulseek.png"), cv2.COLOR_BGR2RGB)
    img_analisis = img.copy()

    # Definir el área de descargas (rectángulo verde)
    descargas_x = seccion_descargas.x
    descargas_y = seccion_descargas.y
    rect_width = 900
    rect_height = 500
    img_h, img_w = img_analisis.shape[:2]
    # Ajuste: límite izquierdo 2px más a la izquierda, derecho 5px más a la derecha
    rect_left = max(0, descargas_x - 2)
    rect_top = max(0, descargas_y + 100)
    rect_right = min(img_w, rect_left + rect_width + 520)
    rect_bottom = min(img_h, rect_top + rect_height + 260)


    # Recortar el área de descargas para el análisis OCR
    area_descargas = img[rect_top:rect_bottom, rect_left:rect_right]
    texto_ocr = pytesseract.image_to_string(area_descargas)
    texto_ocr = texto_ocr.lower()
    ocr_data = pytesseract.image_to_data(area_descargas, output_type=pytesseract.Output.DICT)

    # Asociar búsquedas a texto detectado (buscar nombre de canción como substring ignorando símbolos)
    lineas_ocr = [linea.strip() for linea in texto_ocr.splitlines() if linea.strip()]
    lineas_ocr_norm = [normalizar(linea) for linea in lineas_ocr]

    # Mejorar: buscar coincidencias con mayor robustez (similitud >= 0.8 o substring)
    from difflib import SequenceMatcher
    coincidencias_cajas = []
    for idx_cancion, (cancion, autor) in enumerate(canciones):
        cancion_norm = normalizar(cancion)
        busqueda = f"{cancion} {autor}"
        coincidencias = []
        for i, linea_norm in enumerate(lineas_ocr_norm):
            # Coincidencia robusta: substring o similitud alta
            sim = SequenceMatcher(None, cancion_norm, linea_norm).ratio()
            if (cancion_norm and cancion_norm in linea_norm) or sim >= 0.8:
                coincidencias.append((lineas_ocr[i], sim, i))
        if coincidencias:
            # Elegir la coincidencia con mayor similitud
            best_match = max(coincidencias, key=lambda x: x[1])
            match, score, i = best_match
            descargadas.append((busqueda, match))
            print(f"Descargada: {match} ← búsqueda: {busqueda} (match por nombre de canción, score={score:.2f})")
            # Buscar la caja OCR correspondiente a la línea
            for j in range(len(ocr_data['text'])):
                if ocr_data['text'][j].strip() and ocr_data['text'][j].strip().lower() in match.lower():
                    (left, top, width, height) = (ocr_data['left'][j], ocr_data['top'][j], ocr_data['width'][j], ocr_data['height'][j])
                    # Ajustar coordenadas al área de descargas
                    coincidencias_cajas.append((left + rect_left, top + rect_top, width, height, match))
        else:
            no_descargadas.append(busqueda)
            print(f"No detectada en OCR: {busqueda}")

    # Dibujar borde verde del área de descargas
    cv2.rectangle(img_analisis, (rect_left, rect_top), (rect_right, rect_bottom), (0,255,0), 3)
    # Dibujar cajas rojas sobre coincidencias
    for left, top, width, height, match in coincidencias_cajas:
        cv2.rectangle(img_analisis, (left, top), (left+width, top+height), (0,0,255), 2)
    cv2.imwrite("analisis_descargas.png", cv2.cvtColor(img_analisis, cv2.COLOR_RGB2BGR))

    # Guardar detalles de canciones procesadas al final del log
    with open('soulseek_opciones_log.txt', 'a', encoding='utf-8') as flog:
        flog.write("\n--- Detalles de canciones procesadas ---\n")
        for busqueda, archivo in descargadas:
            flog.write(f"Descargada: {archivo} ← búsqueda: {busqueda}\n")
        for c in no_descargadas:
            flog.write(f"No descargada: {c}\n")

    return descargadas, no_descargadas

# 5. Main
def main():
    ruta_excel = input("Ruta del archivo Excel con canciones y autores: ").strip()
    canciones = leer_excel(ruta_excel)
    print(f"Canciones a buscar: {len(canciones)}")
    print("Asegúrate de tener SoulseekQt abierto y visible en pantalla.")
    if not enfocar_soulseek():
        print("No se pudo encontrar la ventana de SoulseekQt. Abre SoulseekQt y vuelve a intentar.")
        return
    barra_busqueda, seccion_descargas, seccion_buscar, primera_opcion = pedir_posiciones()
    descargadas, no_descargadas = automatizar(canciones, barra_busqueda, seccion_descargas, seccion_buscar, primera_opcion)
    print("\nCanciones procesadas:")
    for busqueda, archivo in descargadas:
        print(f"Descargada: {archivo} ← búsqueda: {busqueda}")
    for c in no_descargadas:
        print(f"No descargada: {c}")

if __name__ == "__main__":
    main()
