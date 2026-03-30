"""
app_web.py  –  Automatización META   (Interfaz Web con NiceGUI)
V 3.1
"""

import asyncio
import glob as glob_module
import json
import logging
import os
import re
import sys
import tempfile
import time
from datetime import datetime
import bcrypt
import psycopg2
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from fastapi import Request
from fastapi.responses import JSONResponse
from nicegui import ui, run, app as _nicegui_app



class LogicaAutomatizacion:

    def __init__(self):
        self.config_file = "config.json"
        self.log_file = f"log_automatizacion_{datetime.now().strftime('%Y%m%d')}.txt"
        self.checkpoint_file = f"checkpoint_{datetime.now().strftime('%Y%m%d_%H%M')}.json"
        self.procesando = False
        self.pausado = False
        self.cancelar = False
        self._csv_df_cache = None       
        self._csv_ruta_cache = None     

        
        self.cb_log = lambda msg: None
        self.cb_progreso = lambda actual, total, columna: None

        self.setup_logging()
        self.config = self.cargar_configuracion()

    # ── LOGGING ──────────────────────────────────────────────────────────────

    def setup_logging(self):
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(self.log_file, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__ + '_web')

    # ── CONFIGURACIÓN ────────────────────────────────────────────────────────

    def cargar_configuracion(self):
        config_default = {
            "sheet_url": "",
            "hoja_default": "COPY/wa JUNIO",
            "fila_inicio": "4",
            "fila_fin": "600",
            "columnas_origen": {
                "copy_it": "N",
                "copy_creativo": "S",
                "copy_conjunto": "X"
            },
            "columnas_destino": {
                "copy_it": {"vis": "O", "alc": "P", "reac": "Q", "clics": "R"},
                "copy_creativo": {"vis": "T", "alc": "U", "reac": "V", "clics": "W"},
                "copy_conjunto": {"vis": "Y", "alc": "Z", "reac": "AA", "clics": "AB"}
            }
        }
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                for key, value in config_default.items():
                    if key not in config:
                        config[key] = value
                return config
            except Exception as e:
                self.logger.error(f"Error cargando configuración: {e}")
        return config_default

    def guardar_configuracion(self, sheet_url, hoja_default, fila_inicio, fila_fin):
        config = {
            "sheet_url": sheet_url.strip(),
            "hoja_default": hoja_default.strip(),
            "fila_inicio": fila_inicio.strip(),
            "fila_fin": fila_fin.strip(),
            "columnas_origen": self.config["columnas_origen"],
            "columnas_destino": self.config["columnas_destino"]
        }
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            self.config = config
            self.logger.info("Configuración guardada exitosamente")
            return True, "Configuración guardada correctamente"
        except Exception as e:
            error_msg = f"No se pudo guardar la configuración: {str(e)}"
            self.logger.error(error_msg)
            return False, error_msg

    # ── CHECKPOINTS ──────────────────────────────────────────────────────────

    def guardar_checkpoint(self, titulos_procesados, exitosos, fallidos, indice_actual, columna_actual=""):
        checkpoint_data = {
            "fecha_hora": datetime.now().isoformat(),
            "titulos_procesados": titulos_procesados,
            "exitosos": exitosos,
            "fallidos": fallidos,
            "indice_actual": indice_actual,
            "columna_actual": columna_actual,
            "config_utilizada": {
                "sheet_url": self.config["sheet_url"],
                "hoja_default": self.config["hoja_default"],
                "fila_inicio": self.config["fila_inicio"],
                "fila_fin": self.config["fila_fin"]
            }
        }
        try:
            with open(self.checkpoint_file, 'w', encoding='utf-8') as f:
                json.dump(checkpoint_data, f, indent=2, ensure_ascii=False)
            self.logger.info(f"Checkpoint guardado: {exitosos} exitosos, {fallidos} fallidos, columna {columna_actual}")
        except Exception as e:
            self.logger.error(f"Error guardando checkpoint: {e}")

    def cargar_checkpoint(self):
        checkpoint_files = glob_module.glob("checkpoint_*.json")
        if not checkpoint_files:
            return None
        checkpoint_files.sort(reverse=True)
        try:
            with open(checkpoint_files[0], 'r', encoding='utf-8') as f:
                checkpoint_data = json.load(f)
            self.logger.info(f"Checkpoint encontrado: {checkpoint_files[0]}")
            return checkpoint_data
        except Exception as e:
            self.logger.error(f"Error cargando checkpoint: {e}")
            return None

    def limpiar_checkpoints_antiguos(self):
        try:
            for file in glob_module.glob("checkpoint_*.json"):
                file_age = time.time() - os.path.getmtime(file)
                if file_age > 7 * 24 * 3600:
                    os.remove(file)
                    self.logger.info(f"Checkpoint antiguo eliminado: {file}")
        except Exception as e:
            self.logger.error(f"Error limpiando checkpoints: {e}")

    # ── GOOGLE SHEETS ─────────────────────────────────────────────────────────

    def extraer_sheet_id(self, url):
        pattern = r'/spreadsheets/d/([a-zA-Z0-9-_]+)'
        match = re.search(pattern, url)
        if match:
            return match.group(1)
        return url.strip()

    def obtener_credenciales(self):
        if getattr(sys, 'frozen', False):
            app_path = os.path.dirname(sys.executable)
        else:
            app_path = os.path.dirname(os.path.abspath(__file__))
        credenciales_path = os.path.join(app_path, 'credenciales.json')
        if not os.path.exists(credenciales_path):
            raise Exception(
                f"No se encuentra el archivo 'credenciales.json' en la carpeta del programa.\n"
                f"Busca en: {app_path}"
            )
        return credenciales_path

    def encontrar_columnas_por_encabezado(self, ws, fila_encabezados=1):
        encabezados = ws.row_values(fila_encabezados)
        columnas_origen = {}
        columnas_destino = {'copy_it': {}, 'copy_creativo': {}, 'copy_conjunto': {}}
        ultimo_tipo_copy = None

        for i, encabezado in enumerate(encabezados):
            if not encabezado:
                continue
            enc = encabezado.lower().strip()
            letra = self.numero_a_letra_columna(i + 1)

            if 'copy' in enc:
                if 'it' in enc:
                    columnas_origen['copy_it'] = {'letra': letra, 'indice': i}
                    ultimo_tipo_copy = 'copy_it'
                    self.logger.info(f"✅ COPY IT → columna {letra}")
                elif 'creativo' in enc or 'creative' in enc:
                    columnas_origen['copy_creativo'] = {'letra': letra, 'indice': i}
                    ultimo_tipo_copy = 'copy_creativo'
                    self.logger.info(f"✅ COPY CREATIVO → columna {letra}")
                elif 'conjunto' in enc or 'set' in enc:
                    columnas_origen['copy_conjunto'] = {'letra': letra, 'indice': i}
                    ultimo_tipo_copy = 'copy_conjunto'
                    self.logger.info(f"✅ COPY CONJUNTO → columna {letra}")
            elif ultimo_tipo_copy:
                if any(k in enc for k in ['vis', 'visual', 'visualizaciones']):
                    if enc in ['vis', 'visualizaciones', 'visuals']:
                        columnas_destino[ultimo_tipo_copy]['vis'] = letra
                elif any(k in enc for k in ['alc', 'alcance', 'reach']):
                    columnas_destino[ultimo_tipo_copy]['alc'] = letra
                elif any(k in enc for k in ['reacc', 'reacciones', 'reactions']):
                    columnas_destino[ultimo_tipo_copy]['reac'] = letra
                elif any(k in enc for k in ['clic', 'click']):
                    columnas_destino[ultimo_tipo_copy]['clics'] = letra

        self.logger.info("="*70 + "\nRESUMEN DE COLUMNAS DETECTADAS:\n" + "="*70)
        for tipo_copy, info in columnas_origen.items():
            self.logger.info(f"\n{tipo_copy.upper().replace('_', ' ')}:")
            self.logger.info(f"   Columna mensaje: {info['letra']}")
            if tipo_copy in columnas_destino:
                m = columnas_destino[tipo_copy]
                self.logger.info(f"  VIS: {m.get('vis','NO ENCONTRADA')} | ALC: {m.get('alc','NO ENCONTRADA')} | REACC: {m.get('reac','NO ENCONTRADA')} | CLICS: {m.get('clics','NO ENCONTRADA')}")
        self.logger.info("="*70)

        return {'origen': columnas_origen, 'destino': columnas_destino}

    def validar_columnas_detectadas(self, columnas_detectadas):
        errores = []
        for tipo in ['copy_it', 'copy_creativo', 'copy_conjunto']:
            if tipo not in columnas_detectadas['origen']:
                errores.append(f"No se encontró la columna {tipo.upper().replace('_', ' ')}")
        for tipo_copy in columnas_detectadas['origen']:
            if tipo_copy in columnas_detectadas['destino']:
                metricas = columnas_detectadas['destino'][tipo_copy]
                for metrica in ['vis', 'alc', 'reac', 'clics']:
                    if metrica not in metricas:
                        errores.append(f"No se encontró {metrica.upper()} para {tipo_copy.upper().replace('_', ' ')}")
        return errores

    def numero_a_letra_columna(self, numero):
        resultado = ""
        while numero > 0:
            numero -= 1
            resultado = chr(numero % 26 + ord('A')) + resultado
            numero //= 26
        return resultado

    def letra_a_numero_columna(self, letra):
        letra = letra.upper()
        num = 0
        for c in letra:
            num = num * 26 + (ord(c) - ord('A') + 1)
        return num

    def conectar_google_sheets(self):
        try:
            credenciales_json = self.obtener_credenciales()
            sheet_key = self.extraer_sheet_id(self.config["sheet_url"])
            scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
            creds = Credentials.from_service_account_file(credenciales_json, scopes=scopes)
            gc = gspread.authorize(creds)
            sh = gc.open_by_key(sheet_key)
            try:
                ws = sh.worksheet(self.config["hoja_default"])
            except gspread.WorksheetNotFound:
                hojas_disponibles = sh.worksheets()
                nombre_buscado = self.config["hoja_default"].strip().lower()
                hoja_encontrada = next(
                    (h for h in hojas_disponibles if h.title.strip().lower() == nombre_buscado), None
                )
                if hoja_encontrada:
                    self.logger.info(f"Hoja encontrada: '{hoja_encontrada.title}'")
                    ws = hoja_encontrada
                else:
                    nombres = [s.title for s in hojas_disponibles]
                    raise Exception(f"No se encontró la hoja '{self.config['hoja_default']}'. Disponibles: {', '.join(nombres)}")
            self.logger.info("Conexión exitosa con Google Sheets")
            return ws
        except Exception as e:
            raise Exception(f"Error al conectar con Google Sheets: {str(e)}")

    # ── BÚSQUEDA CSV ──────────────────────────────────────────────────────────

    def limpiar_texto_para_busqueda(self, texto):
        if not texto:
            return ""
        texto = str(texto).strip()
        for char in ['"', "'", "\u201c", "\u201d", "\u2018", "\u2019", "\u00ab", "\u00bb", "\n", "\r", "\t"]:
            texto = texto.replace(char, "")
        texto = re.sub(r'\s+', ' ', texto)
        return texto.strip()

    def es_mensaje_valido(self, texto):
        if not texto or not str(texto).strip():
            return False
        texto = str(texto).strip()
        contenido_invalido = ['true', 'false', '', 'null', 'none', '#N/A', '#ERROR!', ' ', '\n', '\t']
        if texto.lower() in contenido_invalido:
            return False
        if not texto.replace(' ', '').replace('\n', '').replace('\t', '').replace('\r', ''):
            return False
        texto_sin_sep = texto.replace(',', '').replace('.', '').replace(' ', '').replace('-', '')
        if texto_sin_sep.isdigit():
            return False
        if len(texto) < 15:
            return False
        letras = len(re.findall(r'[a-zA-ZáéíóúüñÁÉÍÓÚÜÑ]', texto))
        total = len(texto.replace(' ', ''))
        if total > 0 and letras / total < 0.5:
            return False
        if len(texto.split()) < 3:
            return False
        for patron in [r'^\d+[\d\s,.-]*$', r'^[\W\d]*$', r'^[^\w\s]*$', r'^\s*$']:
            if re.match(patron, texto):
                return False
        return True

    def buscar_titulo_super_agresivo(self, df, titulo_buscar, columna_copy):
        _col_limpia = '__titulo_limpio__' if '__titulo_limpio__' in df.columns else None
        titulo_limpio = self.limpiar_texto_para_busqueda(titulo_buscar)

        # Estrategia 1: Búsqueda exacta
        r = df[df[columna_copy].str.strip().str.lower() == titulo_limpio.lower()]
        if not r.empty:
            self.logger.info(f"Encontrado con búsqueda exacta: {titulo_buscar[:50]}")
            return r

        # Estrategia 2: Contención original
        r = df[df[columna_copy].str.contains(titulo_buscar, case=False, na=False, regex=False)]
        if not r.empty:
            self.logger.info(f"Encontrado con contención original: {titulo_buscar[:50]}")
            return r

        # Estrategia 3: Contención limpia
        if titulo_limpio:
            r = df[df[columna_copy].str.contains(titulo_limpio, case=False, na=False, regex=False)]
            if not r.empty:
                self.logger.info(f"Encontrado con contención limpia: {titulo_buscar[:50]}")
                return r

        # Estrategia 4: Solo caracteres alfanuméricos
        titulo_solo = ' '.join(''.join(c.lower() for c in titulo_limpio if c.isalnum() or c.isspace()).split())
        for idx, fila_csv in df.iterrows():
            base = fila_csv[_col_limpia] if _col_limpia else self.limpiar_texto_para_busqueda(str(fila_csv[columna_copy]))
            csv_solo = ' '.join(''.join(c.lower() for c in base if c.isalnum() or c.isspace()).split())
            if titulo_solo == csv_solo:
                self.logger.info(f"Encontrado removiendo especiales: {titulo_buscar[:50]}")
                return df.iloc[[idx]]

        # Estrategia 5: Palabras clave
        palabras = titulo_limpio.split()
        if len(palabras) >= 3:
            palabras_clave = [p for p in palabras[:5] if len(p) > 2][:3]
            if palabras_clave:
                patron = '.*'.join(palabras_clave)
                try:
                    r = df[df[columna_copy].str.contains(patron, case=False, na=False, regex=True)]
                    if not r.empty:
                        self.logger.info(f"Encontrado con palabras clave: {titulo_buscar[:50]}")
                        return r
                except Exception:
                    pass

        # Estrategia 6: Contención bidireccional
        for idx, fila_csv in df.iterrows():
            titulo_csv = fila_csv[_col_limpia] if _col_limpia else self.limpiar_texto_para_busqueda(fila_csv[columna_copy])
            if titulo_limpio.lower() in titulo_csv.lower() or titulo_csv.lower() in titulo_limpio.lower():
                self.logger.info(f"Encontrado con contención bidireccional: {titulo_buscar[:50]}")
                return df.iloc[[idx]]

        # Estrategia 7: Similitud por palabras
        palabras_titulo = set(titulo_limpio.lower().split())
        if len(palabras_titulo) >= 2:
            for idx, fila_csv in df.iterrows():
                t_csv = fila_csv[_col_limpia] if _col_limpia else self.limpiar_texto_para_busqueda(fila_csv[columna_copy])
                palabras_csv = set(t_csv.lower().split())
                if len(palabras_csv) >= 2:
                    comunes = palabras_titulo.intersection(palabras_csv)
                    similitud = len(comunes) / len(palabras_titulo.union(palabras_csv))
                    if similitud >= 0.6:
                        self.logger.info(f"Encontrado por similitud ({similitud:.2f}): {titulo_buscar[:50]}")
                        return df.iloc[[idx]]

        # Estrategia 8: Primera y última palabra
        if len(palabras) >= 2:
            pri = palabras[0].lower()
            ult = palabras[-1].lower()
            for idx, fila_csv in df.iterrows():
                t_csv = fila_csv[_col_limpia] if _col_limpia else self.limpiar_texto_para_busqueda(fila_csv[columna_copy])
                p_csv = t_csv.lower().split()
                if len(p_csv) >= 2:
                    if (pri in p_csv[0] or p_csv[0] in pri) and (ult in p_csv[-1] or p_csv[-1] in ult):
                        self.logger.info(f"Encontrado por primera/última palabra: {titulo_buscar[:50]}")
                        return df.iloc[[idx]]

        # Estrategia 9: Longitud similar + palabras clave largas
        long_titulo = len(titulo_limpio)
        for idx, fila_csv in df.iterrows():
            t_csv = fila_csv[_col_limpia] if _col_limpia else self.limpiar_texto_para_busqueda(fila_csv[columna_copy])
            long_csv = len(t_csv)
            if abs(long_titulo - long_csv) <= max(long_titulo, long_csv) * 0.2:
                largas_titulo = [p for p in titulo_limpio.lower().split() if len(p) > 3]
                largas_csv = [p for p in t_csv.lower().split() if len(p) > 3]
                coincidencias = sum(1 for p in largas_titulo if any(p in cp or cp in p for cp in largas_csv))
                if coincidencias >= min(2, len(largas_titulo)):
                    self.logger.info(f"Encontrado por longitud similar: {titulo_buscar[:50]}")
                    return df.iloc[[idx]]

        return pd.DataFrame()

    def buscar_titulo_flexible(self, df, titulo_buscar, columna_copy):
        _col_limpia = '__titulo_limpio__' if '__titulo_limpio__' in df.columns else None

        resultado = self.buscar_titulo_super_agresivo(df, titulo_buscar, columna_copy)
        if not resultado.empty:
            return resultado

        titulo_limpio = self.limpiar_texto_para_busqueda(titulo_buscar)

        r = df[df[columna_copy].str.strip().str.lower() == titulo_limpio.lower()]
        if not r.empty:
            return r

        r = df[df[columna_copy].str.contains(titulo_buscar, case=False, na=False, regex=False)]
        if not r.empty:
            return r

        if titulo_limpio:
            r = df[df[columna_copy].str.contains(titulo_limpio, case=False, na=False, regex=False)]
            if not r.empty:
                return r

        palabras = titulo_limpio.split()
        if len(palabras) >= 3:
            palabras_clave = [p for p in palabras[:5] if len(p) > 2][:3]
            if palabras_clave:
                try:
                    r = df[df[columna_copy].str.contains('.*'.join(palabras_clave), case=False, na=False, regex=True)]
                    if not r.empty:
                        return r
                except Exception:
                    pass

        for idx, fila_csv in df.iterrows():
            titulo_csv = fila_csv[_col_limpia] if _col_limpia else self.limpiar_texto_para_busqueda(fila_csv[columna_copy])
            if titulo_limpio.lower() in titulo_csv.lower() or titulo_csv.lower() in titulo_limpio.lower():
                return df.iloc[[idx]]

        return pd.DataFrame()

    def procesar_csv_por_titulo(self, ruta_archivo, titulo_buscar):
        if not titulo_buscar:
            raise ValueError("El título a buscar no puede estar vacío.")

        # Cargar el CSV solo una vez y reutilizarlo en memoria
        if self._csv_df_cache is None or self._csv_ruta_cache != ruta_archivo:
            df = None
            for encoding in ['utf-8', 'latin1', 'cp1252', 'iso-8859-1']:
                try:
                    df = pd.read_csv(ruta_archivo, encoding=encoding)
                    break
                except UnicodeDecodeError:
                    continue
            if df is None:
                raise Exception("No se pudo leer el archivo con ningún formato de codificación")

            df.columns = df.columns.astype(str)
            columna_copy = "Título"
            if columna_copy not in df.columns:
                raise Exception(f"No se encontró la columna '{columna_copy}' en el CSV. Columnas: {list(df.columns)}")

            # Pre-limpiar todos los títulos una sola vez al cargar 
            df['__titulo_limpio__'] = df[columna_copy].apply(
                lambda x: self.limpiar_texto_para_busqueda(str(x)) if pd.notna(x) else ""
            )
            self._csv_df_cache = df
            self._csv_ruta_cache = ruta_archivo
            self.logger.info(f"CSV cargado en memoria: {len(df)} filas pre-procesadas")

        df = self._csv_df_cache
        columna_copy = "Título"
        fila_encontrada = self.buscar_titulo_flexible(df, titulo_buscar, columna_copy)

        if fila_encontrada.empty:
            self.logger.warning(f"No se encontró: '{titulo_buscar[:50]}...'")
            return None

        datos_fila = fila_encontrada.iloc[0]

        def safe_num(value):
            n = pd.to_numeric(value, errors='coerce')
            return 0 if pd.isna(n) else n

        try:
            vis  = datos_fila['Visualizaciones']
            alc  = datos_fila['Alcance']
            reac = datos_fila['Reacciones, comentarios y veces que se compartió']
            clic = datos_fila['Clics en el enlace']
        except KeyError as e:
            raise KeyError(f"No se encontró la columna {e} en el CSV")

        return {
            'VIS':   int(safe_num(vis)),
            'ALC':   int(safe_num(alc)),
            'REACC': int(safe_num(reac)),
            'CLICS': int(safe_num(clic))
        }

    def extraer_titulos_google_sheets(self, ws):
        fila_inicio = int(self.config["fila_inicio"])
        fila_fin    = int(self.config["fila_fin"])
        titulos_encontrados = []

        self.logger.info("Iniciando detección automática de columnas...")
        columnas_detectadas = self.encontrar_columnas_por_encabezado(ws)

        errores = self.validar_columnas_detectadas(columnas_detectadas)
        if errores:
            msg = "No se pudieron detectar todas las columnas necesarias:\n" + "\n".join(errores)
            msg += "\n\nVerifica que los encabezados contengan:"
            msg += "\n• 'COPY IT', 'COPY CREATIVO', 'COPY CONJUNTO'"
            msg += "\n• 'VIS', 'ALC', 'REACC', 'CLICS' (para cada tipo de COPY)"
            raise Exception(msg)

        range_values = ws.get(f"A{fila_inicio}:AZ{fila_fin}")

        nombres = {'copy_it': 'COPY IT', 'copy_creativo': 'COPY CREATIVO', 'copy_conjunto': 'COPY CONJUNTO'}
        for tipo_copy, info_columna in columnas_detectadas['origen'].items():
            self.logger.info(f"\nProcesando: {nombres[tipo_copy]} (Columna {info_columna['letra']})")
            mensajes_validos = 0
            for i, fila_datos in enumerate(range_values):
                fila_actual = fila_inicio + i
                if len(fila_datos) > info_columna['indice'] and fila_datos[info_columna['indice']]:
                    titulo = str(fila_datos[info_columna['indice']]).strip()
                    if titulo and self.es_mensaje_valido(titulo):
                        titulos_encontrados.append({
                            'titulo': titulo,
                            'fila': fila_actual,
                            'tipo': tipo_copy,
                            'columna_origen': info_columna['letra']
                        })
                        mensajes_validos += 1
            self.logger.info(f"Total mensajes válidos en {nombres[tipo_copy]}: {mensajes_validos}")

        self.actualizar_columnas_config(columnas_detectadas)
        self.logger.info(f"Total de mensajes válidos encontrados: {len(titulos_encontrados)}")
        return titulos_encontrados

    def actualizar_columnas_config(self, columnas_detectadas):
        for tipo_copy, info_columna in columnas_detectadas['origen'].items():
            self.config["columnas_origen"][tipo_copy] = info_columna['letra']
        self.config["columnas_destino"] = columnas_detectadas['destino']

    def colorear_celda_mensaje_no_encontrado(self, ws, fila, tipo_copy):
        try:
            columna_mensaje = self.config["columnas_origen"][tipo_copy]
            formato_verde = {
                "backgroundColor": {"red": 0.6, "green": 1.0, "blue": 0.6},
                "textFormat": {"bold": True, "foregroundColor": {"red": 0.0, "green": 0.0, "blue": 0.0}}
            }
            ws.format(f"{columna_mensaje}{fila}", formato_verde)
            self.logger.info(f"Celda coloreada verde: {columna_mensaje}{fila} ({tipo_copy.upper()})")
            return True
        except Exception as e:
            self.logger.error(f"Error coloreando celda fila {fila}: {str(e)}")
            return False

    def limpiar_celdas_metricas(self, ws, fila, tipo_copy):
        try:
            if tipo_copy not in self.config["columnas_destino"]:
                return False
            columnas = self.config["columnas_destino"][tipo_copy]
            for col_req in ['vis', 'alc', 'reac', 'clics']:
                if col_req not in columnas:
                    return False
            ws.batch_update([
                {'range': f"{columnas['vis']}{fila}",  'values': [[""]]},
                {'range': f"{columnas['alc']}{fila}",  'values': [[""]]},
                {'range': f"{columnas['reac']}{fila}", 'values': [[""]]},
                {'range': f"{columnas['clics']}{fila}", 'values': [[""]]},
            ])
            return True
        except Exception as e:
            self.logger.error(f"Error limpiando celdas fila {fila}: {str(e)}")
            return False

    def subir_datos_a_sheets(self, ws, datos, fila, tipo_copy):
        try:
            if tipo_copy not in self.config["columnas_destino"]:
                raise Exception(f"No se encontraron columnas de destino para {tipo_copy}")
            columnas = self.config["columnas_destino"][tipo_copy]
            for col_req in ['vis', 'alc', 'reac', 'clics']:
                if col_req not in columnas:
                    self.logger.warning(f"Columna {col_req} no encontrada para {tipo_copy}, saltando...")
                    return False

            self.limpiar_celdas_metricas(ws, fila, tipo_copy)
            ws.batch_update([
                {'range': f"{columnas['vis']}{fila}",  'values': [[datos['VIS']]]},
                {'range': f"{columnas['alc']}{fila}",  'values': [[datos['ALC']]]},
                {'range': f"{columnas['reac']}{fila}", 'values': [[datos['REACC']]]},
                {'range': f"{columnas['clics']}{fila}", 'values': [[datos['CLICS']]]},
            ])
            try:
                ws.format(
                    f"{columnas['vis']}{fila}:{columnas['clics']}{fila}",
                    {"numberFormat": {"type": "NUMBER", "pattern": "#,##0"}}
                )
            except Exception as fe:
                self.logger.warning(f"No se pudo aplicar formato fila {fila}: {fe}")
            return True
        except Exception as e:
            self.logger.error(f"Error subiendo datos a fila {fila}: {str(e)}")
            return False

    def validar_entradas(self, sheet_url, hoja, fila_inicio, fila_fin, ruta_csv):
        errores = []
        if not ruta_csv or not str(ruta_csv).strip():
            errores.append("Selecciona un archivo CSV")
        elif not os.path.exists(str(ruta_csv).strip()):
            errores.append("El archivo CSV no existe")
        try:
            self.obtener_credenciales()
        except Exception as e:
            errores.append(str(e))
        if not str(sheet_url).strip():
            errores.append("Ingresa la URL de Google Sheets")
        if not str(hoja).strip():
            errores.append("Ingresa el nombre de la pestaña")
        try:
            fi = int(str(fila_inicio).strip())
            ff = int(str(fila_fin).strip())
            if fi < 1 or ff < 1:
                errores.append("Las filas deben ser números mayores a 0")
            if fi >= ff:
                errores.append("La fila de fin debe ser mayor que la fila de inicio")
        except ValueError:
            errores.append("Las filas de inicio y fin deben ser números válidos")
        return errores

    # ── PROCESAMIENTO CORE ────────────────────────────────────────────────────

    def procesar_core(self, archivo_csv, sheet_url, hoja, fila_inicio, fila_fin, usar_checkpoint=False):
        """
        Función principal de procesamiento.
        Usa cb_log y cb_progreso para actualizar la UI web en tiempo real.
        Retorna: (ok: bool, mensaje: str, resumen: dict | None)
        """
        self.config["sheet_url"]    = str(sheet_url).strip()
        self.config["hoja_default"] = str(hoja).strip()
        self.config["fila_inicio"]  = str(fila_inicio).strip()
        self.config["fila_fin"]     = str(fila_fin).strip()

        errores = self.validar_entradas(sheet_url, hoja, fila_inicio, fila_fin, archivo_csv)
        if errores:
            return False, "Errores de validación:\n" + "\n".join(errores), None

        checkpoint = self.cargar_checkpoint() if usar_checkpoint else None

        try:
            self.cb_log("Conectando con Google Sheets...")
            ws = self.conectar_google_sheets()

            self.cb_log(f"Extrayendo títulos (filas {self.config['fila_inicio']}-{self.config['fila_fin']})...")
            titulos = self.extraer_titulos_google_sheets(ws)

            if not titulos:
                return False, f"No se encontraron títulos en el rango {self.config['fila_inicio']}-{self.config['fila_fin']}", None

            if checkpoint and usar_checkpoint:
                indice_inicio = checkpoint["indice_actual"]
                exitosos      = checkpoint["exitosos"]
                fallidos      = checkpoint["fallidos"]
                self.cb_log(f"Continuando desde elemento {indice_inicio + 1}/{len(titulos)}")
            else:
                indice_inicio = 0
                exitosos      = 0
                fallidos      = 0
                self.limpiar_checkpoints_antiguos()

            self.cb_log(f"Total a procesar: {len(titulos)} títulos")

            columna_actual    = ""
            exitosos_columna  = 0
            total_columna     = 0
            resultados_columnas = {}

            for i in range(indice_inicio, len(titulos)):
                if self.cancelar:
                    self.logger.info("Procesamiento cancelado por el usuario")
                    break

                while self.pausado and not self.cancelar:
                    time.sleep(0.5)

                if self.cancelar:
                    break

                titulo_info = titulos[i]

                if titulo_info['tipo'] != columna_actual:
                    if columna_actual and total_columna > 0:
                        resultados_columnas[columna_actual] = (exitosos_columna, total_columna)
                    columna_actual   = titulo_info['tipo']
                    exitosos_columna = 0
                    total_columna    = 0
                    self.logger.info(f"Iniciando procesamiento de {columna_actual.upper()}")

                total_columna += 1
                self.cb_log(f"[{i+1}/{len(titulos)}] {titulo_info['titulo'][:65]}...")
                self.cb_progreso(i + 1, len(titulos), titulo_info['tipo'])

                try:
                    datos = self.procesar_csv_por_titulo(archivo_csv, titulo_info['titulo'])
                    if datos:
                        if self.subir_datos_a_sheets(ws, datos, titulo_info['fila'], titulo_info['tipo']):
                            exitosos         += 1
                            exitosos_columna += 1
                            self.cb_log(f"Exitoso · Fila {titulo_info['fila']} ({titulo_info['tipo'].upper().replace('_', ' ')})")
                        else:
                            fallidos += 1
                            self.cb_log(f"Error al subir · Fila {titulo_info['fila']}")
                    else:
                        if self.colorear_celda_mensaje_no_encontrado(ws, titulo_info['fila'], titulo_info['tipo']):
                            fallidos += 1
                            self.cb_log(f"No encontrado · Celda coloreada rojo · Fila {titulo_info['fila']}")
                        else:
                            fallidos += 1
                            self.cb_log(f"No encontrado en CSV · Fila {titulo_info['fila']}")

                    time.sleep(1.5)

                except Exception as e:
                    fallidos += 1
                    self.cb_log(f"  ❌ Error: {str(e)[:80]}")
                    self.logger.error(f"Error procesando '{titulo_info['titulo']}': {e}")

                if (i + 1) % 3  == 0: time.sleep(2.0)
                if (i + 1) % 15 == 0: time.sleep(5.0)
                if (i + 1) % 10 == 0:
                    self.guardar_checkpoint(i + 1, exitosos, fallidos, i + 1, columna_actual)
                if (i + 1) % 30  == 0: time.sleep(10.0)
                if (i + 1) % 100 == 0:
                    time.sleep(30.0)
                    self.guardar_checkpoint(i + 1, exitosos, fallidos, i + 1, columna_actual)

            self.procesando = False

            if self.cancelar:
                try:
                    self.guardar_checkpoint(i + 1, exitosos, fallidos, i + 1, columna_actual)
                except Exception:
                    self.guardar_checkpoint(len(titulos), exitosos, fallidos, len(titulos), columna_actual)
                return False, "Procesamiento cancelado por el usuario", None

            if columna_actual and total_columna > 0:
                resultados_columnas[columna_actual] = (exitosos_columna, total_columna)

            self.guardar_checkpoint(len(titulos), exitosos, fallidos, len(titulos), "COMPLETADO")

            nombres_cols = {
                "copy_it": "COPY IT",
                "copy_creativo": "COPY CREATIVO",
                "copy_conjunto": "COPY CONJUNTO"
            }
            detalle = {}
            for col, (ex, tot) in resultados_columnas.items():
                prec = (ex / tot * 100) if tot > 0 else 0
                detalle[nombres_cols.get(col, col.upper())] = {"exitosos": ex, "total": tot, "precision": prec}

            resumen = {
                "titulos_total": len(titulos),
                "exitosos":  exitosos,
                "fallidos":  fallidos,
                "precision": (exitosos / len(titulos) * 100) if titulos else 0,
                "fila_inicio": self.config["fila_inicio"],
                "fila_fin":    self.config["fila_fin"],
                "detalle_columnas": detalle
            }

            self.logger.info(
                f"Procesamiento completado: {exitosos} exitosos, {fallidos} fallidos, "
                f"{resumen['precision']:.1f}% precisión"
            )
            return True, "Procesamiento completado exitosamente", resumen

        except Exception as e:
            self.procesando = False
            error_msg = f"Error durante el procesamiento: {str(e)}"
            self.logger.error(error_msg)
            return False, error_msg, None


# ──────────────────────────────────────────────────────────────────────────────
#  INTERFAZ WEB  (NiceGUI)
# ──────────────────────────────────────────────────────────────────────────────

logica = LogicaAutomatizacion()
_state = {'csv_path': None, 'pending_name': None}  # sin problemas de scope en closures

# Servir archivos estáticos (imágenes, logos)
_nicegui_app.add_static_files('/static', os.path.dirname(os.path.abspath(__file__)))


@_nicegui_app.post('/api/upload-csv')
async def _api_upload_csv(request: Request):
    """Recibe el CSV subido desde el navegador y actualiza el estado."""
    form = await request.form()
    file_obj = form.get('file')
    if not file_obj:
        return JSONResponse({'ok': False, 'error': 'No file'}, status_code=400)
    content = await file_obj.read()
    filename = file_obj.filename or 'archivo.csv'
    suffix = os.path.splitext(filename)[1] or '.csv'
    app_dir = os.path.dirname(os.path.abspath(__file__))
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix, dir=app_dir)
    tmp.write(content)
    tmp.close()
    _state['csv_path'] = tmp.name
    _state['pending_name'] = filename
    logica._csv_df_cache   = None
    logica._csv_ruta_cache = None
    return JSONResponse({'ok': True, 'name': filename})


# ──────────────────────────────────────────────────────────────────────────────
#  PÁGINA DE LOGIN
# ──────────────────────────────────────────────────────────────────────────────

@ui.page('/login')
def login_page():
    # Si ya está autenticado, redirigir directo a la app
    if _nicegui_app.storage.user.get('autenticado'):
        ui.navigate.to('/')
        return

    ui.add_head_html("""
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
    <style>
      *, *::before, *::after { box-sizing: border-box; font-family: 'Inter', sans-serif; margin: 0; padding: 0; }
      /* Fondo oscuro de fallback (si el video no carga) */
      html, body { height: 100%; background: #0a0f1e; }

      /* ── Todos los wrappers de Quasar/NiceGUI transparentes ── */
      #app, .q-layout, .q-layout__shadow,
      .q-page-container, .q-page,
      body.body--dark, body.body--dark .q-layout,
      body.body--dark .q-page-container {
        background: transparent !important;
      }

      /* ── Video de fondo (inyectado en body via JS) ── */
      #login-video-bg {
        position: fixed; top: 0; left: 0;
        width: 100%; height: 100%;
        object-fit: cover;
        z-index: 0;
        pointer-events: none;
        /* Empieza borroso, se aclara al terminar el splash */
        filter: blur(18px);
        transform: scale(1.08);
        transition: filter 1.6s ease, transform 1.6s ease;
      }
      #login-video-bg.unblurred {
        filter: blur(0px);
        transform: scale(1);
      }
      /* ── Overlay oscuro sobre el video ── */
      #login-video-overlay {
        position: fixed; top: 0; left: 0;
        width: 100%; height: 100%;
        background: rgba(8, 12, 24, 0.40);
        z-index: 1;
        pointer-events: none;
      }

      /* ── Centrar contenido de la página NiceGUI ── */
      .q-page {
        display: flex !important;
        flex-direction: column;
        align-items: center !important;
        justify-content: center !important;
        min-height: 100vh !important;
        position: relative;
        z-index: 2;
      }

      /* ── Orbs decorativos ── */
      .orb {
        position: fixed; border-radius: 50%; pointer-events: none;
        animation: orb-drift 18s ease-in-out infinite;
      }
      .orb-1 {
        width: 620px; height: 620px;
        background: radial-gradient(circle, rgba(99,102,241,0.10) 0%, transparent 70%);
        top: -260px; right: -260px;
      }
      .orb-2 {
        width: 500px; height: 500px;
        background: radial-gradient(circle, rgba(139,92,246,0.08) 0%, transparent 70%);
        bottom: -220px; left: -220px;
        animation-duration: 24s; animation-direction: reverse;
      }
      @keyframes orb-drift {
        0%,100% { transform: translate(0,0) scale(1); }
        33%      { transform: translate(28px,-28px) scale(1.08); }
        66%      { transform: translate(-20px,18px) scale(0.93); }
      }

      /* ── Card base – glassmorphism ── */
      .auth-card {
        background: rgba(10, 16, 35, 0.38);
        backdrop-filter: blur(32px) saturate(160%);
        -webkit-backdrop-filter: blur(32px) saturate(160%);
        border: 1px solid rgba(255,255,255,0.18);
        border-radius: 26px;
        padding: 80px 48px;
        width: 100%; max-width: 520px;
        min-height: 680px;
        display: flex; flex-direction: column; justify-content: center;
        box-shadow: 0 24px 60px rgba(0,0,0,0.45),
                    inset 0 1px 0 rgba(255,255,255,0.12);
        position: relative; z-index: 10;
      }

      /* ── Logo flotante ── */
      .logo-img {
        display: block; width: 140px; height: auto;
        max-height: 120px; object-fit: contain;
        margin: 0 auto 18px auto;
        animation: logo-float 3.2s ease-in-out infinite;
        filter: drop-shadow(0 6px 24px rgba(99,102,241,0.5));
      }
      @keyframes logo-float {
        0%,100% { transform: translateY(0); }
        50%      { transform: translateY(-9px); }
      }
      .brand-title {
        font-size: 1.35rem; font-weight: 700; color: #ffffff;
        text-align: center; letter-spacing: -0.3px;
        text-shadow: 0 2px 12px rgba(0,0,0,0.9), 0 1px 3px rgba(0,0,0,0.7);
      }
      .brand-sub {
        font-size: 0.76rem; color: #e2e8f0; text-align: center;
        text-transform: uppercase; letter-spacing: 1.1px; margin-top: 5px;
        text-shadow: 0 1px 6px rgba(0,0,0,0.8);
      }

      /* ── Dots de carga ── */
      .dots { display: flex; gap: 7px; justify-content: center; margin-top: 28px; }
      .dot {
        width: 9px; height: 9px; background: #6366f1; border-radius: 50%;
        animation: dot-bounce 1.4s ease-in-out infinite;
      }
      .dot:nth-child(2) { animation-delay: 0.2s; }
      .dot:nth-child(3) { animation-delay: 0.4s; }
      @keyframes dot-bounce {
        0%,80%,100% { transform: scale(0.55); opacity: 0.35; }
        40%          { transform: scale(1); opacity: 1; }
      }
      .loading-txt {
        text-align: center; color: #cbd5e1; font-size: 0.82rem; margin-top: 14px;
        text-shadow: 0 1px 6px rgba(0,0,0,0.8);
        animation: txt-pulse 2s ease-in-out infinite;
      }
      @keyframes txt-pulse { 0%,100% { opacity:0.45; } 50% { opacity:1; } }

      /* ── Splash card ── */
      #splash-card { animation: card-pop 0.65s cubic-bezier(0.34,1.56,0.64,1) forwards; }
      @keyframes card-pop {
        from { opacity:0; transform: scale(0.88) translateY(22px); }
        to   { opacity:1; transform: scale(1)    translateY(0);    }
      }
      .auth-hr { border:none; border-top:1px solid rgba(255,255,255,0.15); margin:28px 0; }

      /* ── Login form card ── */
      ._login_form {
        display: none;
        flex-direction: column; align-items: center; justify-content: center;
        z-index: 10; position: relative;
        width: 100%; max-width: 520px;
      }
      ._login_inner {
        width: 100%;
        background: rgba(10, 16, 35, 0.38);
        backdrop-filter: blur(32px) saturate(160%);
        -webkit-backdrop-filter: blur(32px) saturate(160%);
        border: 1px solid rgba(255,255,255,0.18);
        border-radius: 26px;
        padding: 80px 48px;
        min-height: 680px;
        display: flex; flex-direction: column; justify-content: center;
        box-shadow: 0 24px 60px rgba(0,0,0,0.45),
                    inset 0 1px 0 rgba(255,255,255,0.12);
        animation: card-pop 0.6s cubic-bezier(0.34,1.56,0.64,1) forwards;
      }
      @keyframes form-shake {
        0%,100% { transform: translateX(0); }
        20%,60% { transform: translateX(-7px); }
        40%,80% { transform: translateX(7px); }
      }
      .shake { animation: form-shake 0.45s ease !important; }

      /* ── Inputs glass ── */
      ._login_form .q-field--outlined .q-field__control {
        border-radius: 14px !important;
        background: rgba(255,255,255,0.07) !important;
        border-color: rgba(255,255,255,0.22) !important;
      }
      ._login_form .q-field--outlined.q-field--focused .q-field__control {
        background: rgba(255,255,255,0.11) !important;
        border-color: rgba(99,102,241,0.7) !important;
      }
      ._login_form .q-field__native,
      ._login_form .q-field__input {
        color: #ffffff !important;
      }
      /* ── Texto con padding interno, ojo separado del borde ── */
      ._login_form .q-field__native,
      ._login_form .q-field__input {
        padding-left: 16px !important;
        padding-right: 8px !important;
      }
      ._login_form .q-field__append {
        padding-right: 12px !important;
      }

      /* ── Placeholder: color base ── */
      ._login_form .q-field__native::placeholder,
      ._login_form .q-field__input::placeholder {
        color: rgba(148, 163, 184, 0.65);
        transition: opacity 0.25s;
      }
      /* Solo el campo Usuario desaparece al hacer clic */
      ._login_form .inp-usuario .q-field__native:focus::placeholder,
      ._login_form .inp-usuario .q-field__input:focus::placeholder {
        opacity: 0;
      }
      /* Ocultar etiqueta flotante de Quasar en el login */
      ._login_form .q-field__label {
        display: none !important;
      }

      /* ── Success overlay ── */
      #success-overlay {
        position: fixed; inset: 0;
        background: rgba(8, 12, 24, 0.80);
        display: none; flex-direction: column;
        align-items: center; justify-content: center;
        gap: 18px; z-index: 9999;
        backdrop-filter: blur(2px);
      }
      #success-overlay.show { display: flex; animation: fade-in 0.4s ease; }
      @keyframes fade-in { from { opacity:0; } to { opacity:1; } }
      .success-ring {
        width: 110px; height: 110px;
        background: rgba(99,102,241,0.10);
        border: 2px solid rgba(99,102,241,0.4); border-radius: 50%;
        display: flex; align-items: center; justify-content: center;
        padding: 14px;
        animation: ring-pop 0.55s cubic-bezier(0.34,1.56,0.64,1);
      }
      @keyframes ring-pop {
        from { transform: scale(0); opacity:0; }
        to   { transform: scale(1); opacity:1; }
      }
    </style>
    """)

    # ── Video de fondo inyectado directo en <body> para evitar wrappers Quasar ──
    ui.add_head_html("""
    <script>
    (function() {
      function injectVideoBg() {
        if (document.getElementById('login-video-bg')) return;
        var v = document.createElement('video');
        v.id = 'login-video-bg';
        v.autoplay = true; v.loop = true; v.muted = true;
        v.setAttribute('playsinline', '');
        var s = document.createElement('source');
        s.src = '/static/login_video_2.mp4'; s.type = 'video/mp4';
        v.appendChild(s);
        document.body.insertBefore(v, document.body.firstChild);
        var o = document.createElement('div');
        o.id = 'login-video-overlay';
        document.body.insertBefore(o, v.nextSibling);
        v.play().catch(function(){});
      }
      if (document.body) { injectVideoBg(); }
      else { document.addEventListener('DOMContentLoaded', injectVideoBg); }
    })();
    </script>
    """)

    # ── Splash (Phase 1) ──────────────────────────────────────────────────────
    ui.html("""
    <div id="splash-card" class="auth-card" style="text-align:center;">
      <img src="/static/logo_valv.png" class="logo-img" alt="VALV">
      <div class="brand-title">Bienvenid@ a Vuela a la Vida</div>
      <div class="brand-sub">Sistema Interno</div>
      <hr class="auth-hr">
      <div class="dots">
        <div class="dot"></div><div class="dot"></div><div class="dot"></div>
      </div>
      <div class="loading-txt">Iniciando sistema...</div>
    </div>
    """)

    # ── Login form (Phase 2) – NiceGUI elements, hidden initially ────────────
    with ui.element('div').classes('_login_form'):
        with ui.element('div').classes('_login_inner'):

            # Logo + título
            ui.html("""
            <div style="text-align:center; margin-bottom:10px;">
              <img src="/static/logo_valv.png" class="logo-img" alt="VALV">
              <div class="brand-title" style="font-size:1.3rem; margin-top:6px;">Bienvenid@</div>
              <div class="brand-sub" style="color:#e2e8f0; margin-top:8px;">Ingresa tus credenciales</div>
            </div>
            <hr class="auth-hr">
            """)

            # Mensaje de error (oculto hasta que falle)
            error_label = ui.label('').style(
                'color:#f87171; font-size:0.82rem; text-align:center; '
                'min-height:20px; display:block; margin-bottom:12px;'
            )

            # Campos
            inp_user = ui.input(
                placeholder='Usuario'
            ).props('outlined color=indigo').classes('w-full inp-usuario')

            inp_pass = ui.input(
                placeholder='••••••••', password=True, password_toggle_button=True,
            ).props('outlined color=indigo').classes('w-full').style('margin-top:20px')

            # Handler de login
            async def do_login():
                if not inp_user.value.strip() or not inp_pass.value.strip():
                    error_label.set_text('Por favor completa todos los campos')
                    await ui.run_javascript(
                        "document.querySelector('._login_inner').classList.add('shake');"
                        "setTimeout(()=>document.querySelector('._login_inner').classList.remove('shake'),500);"
                    )
                    return

                # Validar contra PostgreSQL
                email = inp_user.value.strip().lower()
                password = inp_pass.value.strip()
                usuario_valido = False
                nombre_usuario = ''
                try:
                    conn = psycopg2.connect(
                        host=os.environ.get('DB_HOST', 'localhost'),
                        database=os.environ.get('DB_NAME', 'meta_valv'),
                        user=os.environ.get('DB_USER', 'valv_user'),
                        password=os.environ.get('DB_PASSWORD', ''),
                        port=int(os.environ.get('DB_PORT', '5432'))
                    )
                    cur = conn.cursor()
                    cur.execute(
                        "SELECT password_hash, nombre, activo FROM usuarios WHERE email = %s",
                        (email,)
                    )
                    row = cur.fetchone()
                    cur.close()
                    conn.close()
                    if row and row[2] and bcrypt.checkpw(password.encode('utf-8'), row[0].encode('utf-8')):
                        usuario_valido = True
                        nombre_usuario = row[1]
                except Exception as e:
                    logging.error(f'Error DB login: {e}')

                if not usuario_valido:
                    error_label.set_text('Correo o contraseña incorrectos')
                    await ui.run_javascript(
                        "document.querySelector('._login_inner').classList.add('shake');"
                        "setTimeout(()=>document.querySelector('._login_inner').classList.remove('shake'),500);"
                    )
                    return

                _nicegui_app.storage.user['autenticado'] = True
                _nicegui_app.storage.user['usuario'] = nombre_usuario
                await ui.run_javascript(
                    "document.getElementById('success-overlay').classList.add('show');"
                )
                await asyncio.sleep(2.0)
                ui.navigate.to('/')

            inp_pass.on('keydown.enter', do_login)
            ui.button('Ingresar', on_click=do_login).props(
                'unelevated color=indigo no-caps'
            ).classes('w-full').style(
                'height:52px; font-size:1rem; font-weight:600; '
                'border-radius:14px; letter-spacing:0.3px; margin-top:28px;'
            )

            ui.html("""
            <div style="display:flex;align-items:center;justify-content:center;
                        gap:5px;margin-top:26px;color:#e2e8f0;font-size:0.74rem;text-shadow:0 1px 6px rgba(0,0,0,0.8);">
              <span class="material-icons" style="font-size:0.9rem;">lock</span>
              Acceso restringido personal autorizado
            </div>
            """)

    # ── Success overlay (Phase 3) ─────────────────────────────────────────────
    ui.html("""
    <div id="success-overlay">
      <div class="success-ring">
        <img src="/static/logo_valv.png" style="width:100%;height:100%;object-fit:contain;" alt="VALV">
      </div>
      <div style="color:#f1f5f9;font-size:1.2rem;font-weight:700;display:flex;align-items:center;gap:8px;">
        ✅ Acceso autorizado
      </div>
      <div style="color:#94a3b8;font-size:0.92rem;font-weight:500;">Cargando tu espacio de trabajo...</div>
      <div class="dots" style="margin-top:4px;">
        <div class="dot"></div><div class="dot"></div><div class="dot"></div>
      </div>
    </div>
    """)

    # ── JS: transición splash → formulario ───────────────────────────────────
    ui.add_head_html("""
    <script>
    setTimeout(function() {
      var splash = document.getElementById('splash-card');
      if (!splash) return;
      // Desvanecer splash
      splash.style.transition = 'opacity 0.5s, transform 0.5s';
      splash.style.opacity = '0';
      splash.style.transform = 'scale(0.94) translateY(-10px)';
      // Desenfocar el video al mismo tiempo
      var video = document.getElementById('login-video-bg');
      if (video) { video.classList.add('unblurred'); }
      setTimeout(function() {
        splash.style.display = 'none';
        var form = document.querySelector('._login_form');
        if (form) { form.style.display = 'flex'; }
      }, 520);
    }, 2500);
    </script>
    """)


@ui.page('/')
def index():
    if not _nicegui_app.storage.user.get('autenticado'):
        ui.navigate.to('/login')
        return

    # ── HEAD: fuentes + estilos ───────────────────────────────────────────
    ui.add_head_html("""
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
    <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
    <style>
      * { font-family: 'Inter', sans-serif; }
      .mono { font-family: 'JetBrains Mono', monospace !important; }

      /* ── Cards redondeadas y con borde sutil ── */
      .app-card {
        border-radius: 22px !important;
        margin-bottom: 16px;
        overflow: visible !important;
        transition: box-shadow 0.2s, border-color 0.2s;
      }
      body.body--dark .app-card {
        border: 1px solid rgba(255,255,255,0.07) !important;
        box-shadow: 0 4px 28px rgba(0,0,0,0.40) !important;
      }
      body.body--light .app-card {
        border: 1px solid rgba(0,0,0,0.07) !important;
        box-shadow: 0 4px 20px rgba(0,0,0,0.07) !important;
      }

      /* ── Título de sección: contraste en ambos modos ── */
      .section-title {
        font-size: 1.05rem; font-weight: 700;
        margin-bottom: 14px; letter-spacing: -0.2px;
      }
      body.body--dark  .section-title { color: #f1f5f9; }
      body.body--light .section-title { color: #0f172a; }

      /* ── Todos los inputs: bordes redondeados ── */
      .q-field--outlined .q-field__control {
        border-radius: 14px !important;
      }
      /* Modo oscuro – inputs */
      body.body--dark .q-field--outlined .q-field__control {
        background: rgba(255,255,255,0.04) !important;
        border-color: rgba(255,255,255,0.14) !important;
      }
      body.body--dark .q-field--outlined.q-field--focused .q-field__control {
        border-color: #6366f1 !important;
        background: rgba(99,102,241,0.06) !important;
      }
      body.body--dark .q-field__native,
      body.body--dark .q-field__input  { color: #f1f5f9 !important; }
      body.body--dark .q-field__label  { color: #94a3b8 !important; }

      /* Modo claro – inputs */
      body.body--light .q-field--outlined .q-field__control {
        background: rgba(0,0,0,0.025) !important;
        border-color: rgba(0,0,0,0.18) !important;
      }
      body.body--light .q-field--outlined.q-field--focused .q-field__control {
        border-color: #6366f1 !important;
        background: rgba(99,102,241,0.04) !important;
      }
      body.body--light .q-field__native,
      body.body--light .q-field__input  { color: #0f172a !important; }
      body.body--light .q-field__label  { color: #475569 !important; }

      /* ── Botones: esquinas redondeadas ── */
      .q-btn:not(.q-btn--round) { border-radius: 12px !important; }

      /* ── Dialogs/modals redondeados ── */
      .q-dialog__inner > .q-card { border-radius: 22px !important; }

      /* ── Texto general: contraste por modo ── */
      body.body--dark  .q-card  { color: #e2e8f0; }
      body.body--light .q-card  { color: #1e293b; }
      body.body--dark  .text-grey-5 { color: #94a3b8 !important; }
      body.body--light .text-grey-5 { color: #475569 !important; }

      /* ── Stat chips ── */
      .stat-chip {
        border-radius: 18px; padding: 12px 20px;
        text-align: center; min-width: 120px;
      }
      /* ── Badge ── */
      .badge {
        display: inline-flex; align-items: center; gap: 6px;
        padding: 4px 10px; border-radius: 20px;
        font-size: 0.78rem; font-weight: 500;
      }
      /* ── Progress bar ── */
      .q-linear-progress { border-radius: 10px; overflow: hidden; }
      .q-linear-progress__track,
      .q-linear-progress__model { border-radius: 10px !important; }

      /* ── Feature row ── */
      .feature-row { display:flex; align-items:center; gap:8px; margin:4px 0; font-size:0.85rem; }

      /* ── CSV drop zone ── */
      .csv-dropzone {
        cursor: pointer;
        border: 2px dashed rgba(99,102,241,0.55);
        border-radius: 16px;
        padding: 20px 24px;
        text-align: center;
        transition: background 0.25s, border-color 0.25s, transform 0.15s;
        width: 100%;
        display: flex; align-items: center; justify-content: center; gap: 10px;
      }
      .csv-dropzone:hover {
        background: rgba(99,102,241,0.08);
        border-color: #818cf8;
        transform: translateY(-1px);
      }

      /* ── Selector de tema (toggle interno) ── */
      .q-toggle__inner { border-radius: 999px !important; }

      /* ── Notify / toasts redondeados ── */
      .q-notification { border-radius: 14px !important; }
    </style>
    """)

    # ── CSV upload JS──
    ui.add_head_html("""
    <script>
    var _csvSetup = setInterval(function() {
        var mount = document.getElementById('csv-mount');
        if (!mount || mount._done) return;
        mount._done = true;
        clearInterval(_csvSetup);

        var inp = document.createElement('input');
        inp.type = 'file';
        inp.accept = '.csv';
        inp.style.display = 'none';
        inp.addEventListener('change', async function(e) {
            var file = e.target.files[0];
            if (!file) return;
            var fd = new FormData();
            fd.append('file', file);
            try { await fetch('/api/upload-csv', {method:'POST', body:fd}); }
            catch(err) { console.error('CSV upload error:', err); }
            e.target.value = '';
        });

        var lbl = document.createElement('label');
        lbl.className = 'csv-dropzone';
        lbl.style.cssText = 'width:100%;box-sizing:border-box;cursor:pointer;';
        lbl.innerHTML =
            '<span class="material-icons" style="font-size:1.8rem;color:#6366f1;pointer-events:none">upload_file</span>'
          + '<span style="font-size:.95rem;font-weight:600;color:#6366f1;pointer-events:none">'
          + '📂\u00a0\u00a0Seleccionar Archivo CSV</span>';
        lbl.appendChild(inp);
        mount.appendChild(lbl);
    }, 50);
    </script>
    """)

    dark_mode = ui.dark_mode(value=True)

    with ui.column().classes('w-full items-center q-pa-md'):
        with ui.column().classes('w-full').style('max-width: 980px'):

            # ── HEADER ───────────────────────────────────────────────────
            with ui.card().classes('app-card w-full q-pa-lg'):
                with ui.row().classes('items-start justify-between w-full flex-wrap gap-4'):
                    with ui.column().classes('gap-1'):
                        with ui.row().classes('items-center gap-3'):
                            ui.image('/static/logo_valv.png').style(
                                'width:72px; height:auto; object-fit:contain; display:block; flex-shrink:0;'
                            )
                            ui.label('Automatización-META SMM').style(
                                'font-size:2rem; font-weight:700; letter-spacing:-0.5px;'
                            )
                        ui.label('Detección automática de columnas · Procesamiento inteligente').classes('text-grey-5').style('font-size:0.85rem')

                    with ui.row().classes('items-center gap-3'):
                        with ui.card().classes('q-pa-sm').style('border-radius:18px; min-width:220px'):
                            ui.label('Tema de la interfaz').classes('text-grey-5').style('font-size:0.75rem; margin-bottom:6px')
                            ui.toggle(
                                {'dark': '🌙 Oscuro', 'light': '☀️ Claro', 'auto': '🖥️ Sistema'},
                                value='dark',
                                on_change=lambda e: (
                                    dark_mode.enable()  if e.value == 'dark'  else
                                    dark_mode.disable() if e.value == 'light' else
                                    dark_mode.auto()
                                )
                            ).props('dense')

                        async def cerrar_sesion():
                            _nicegui_app.storage.user.clear()
                            ui.navigate.to('/login')

                        ui.button(icon='logout', on_click=cerrar_sesion).props(
                            'flat round dense color=red'
                        ).tooltip('Cerrar sesión')

            # ── ARCHIVO CSV ───────────────────────────────────────────────
            with ui.card().classes('app-card w-full q-pa-lg'):
                ui.label('📁 Archivo de Datos CSV').classes('section-title')

                # Fila de estado: nombre del archivo + botón quitar
                with ui.row().classes('items-center gap-2 q-mb-sm'):
                    csv_status = ui.label('Sin archivo seleccionado').style(
                        'font-size:0.85rem; color:#64748b;'
                    )
                    btn_clear_csv = ui.button(
                        '❌', on_click=lambda: _clear_csv()
                    ).props('flat dense round color=red size=xs no-caps').tooltip('Quitar archivo')
                    btn_clear_csv.set_visibility(False)

                def _clear_csv():
                    _state['csv_path'] = None
                    _state['pending_name'] = None
                    logica._csv_df_cache   = None
                    logica._csv_ruta_cache = None
                    csv_status.set_text('Sin archivo seleccionado')
                    csv_status.style('font-size:0.85rem; color:#64748b;')
                    btn_clear_csv.set_visibility(False)
                    ui.notify('Archivo CSV eliminado', type='info', position='top-right')

                # Timer que detecta cuando el endpoint subió un archivo y actualiza la UI
                def _check_pending():
                    name = _state.get('pending_name')
                    if name:
                        _state['pending_name'] = None
                        csv_status.set_text(f'✅  {name}')
                        csv_status.style('font-size:0.85rem; color:#4ade80;')
                        btn_clear_csv.set_visibility(True)
                        ui.notify(f'CSV cargado: {name}', type='positive', position='top-right')

                ui.timer(0.4, _check_pending)

                # Punto de montaje –
                ui.html('<div id="csv-mount"></div>')

            # ── GOOGLE SHEETS ─────────────────────────────────────────────
            with ui.card().classes('app-card w-full q-pa-lg'):
                ui.label('📊 Configuración de Google Sheets').classes('section-title')

                inp_url = ui.input(
                    label='URL de Google Sheets',
                    placeholder='https://docs.google.com/spreadsheets/d/...',
                    value=logica.config.get('sheet_url', '')
                ).classes('w-full q-mb-md').props('outlined dense')

                with ui.row().classes('w-full gap-4 flex-wrap'):
                    inp_hoja = ui.input(
                        label='Nombre de la pestaña',
                        placeholder='COPY/wa JUNIO',
                        value=logica.config.get('hoja_default', 'COPY/wa JUNIO')
                    ).classes('flex-grow').props('outlined dense')

                    inp_fi = ui.input(
                        label='Fila inicio',
                        placeholder='4',
                        value=logica.config.get('fila_inicio', '4')
                    ).classes('w-28').props('outlined dense')

                    inp_ff = ui.input(
                        label='Fila fin',
                        placeholder='600',
                        value=logica.config.get('fila_fin', '600')
                    ).classes('w-28').props('outlined dense')

            # ── CONTROLES ─────────────────────────────────────────────────
            with ui.card().classes('app-card w-full q-pa-lg'):
                ui.label('🎛️ Controles').classes('section-title')
                with ui.row().classes('justify-center gap-4 flex-wrap w-full'):

                    BTN_STYLE = 'min-width: 160px; height: 44px; font-size: 0.88rem; font-weight: 600; border-radius: 12px;'

                    # Guardar configuraciones csv mas filas y hoja
                    def guardar_config():
                        ok, msg = logica.guardar_configuracion(
                            inp_url.value, inp_hoja.value, inp_fi.value, inp_ff.value
                        )
                        ui.notify(msg, type='positive' if ok else 'negative', position='top-right')

                    ui.button('💾 Guardar Config', on_click=guardar_config).props(
                        'unelevated color=blue no-caps'
                    ).style(BTN_STYLE)

                    # ── Procesar ──────────────────────────────────────────
                    async def iniciar_procesamiento():
                        if logica.procesando:
                            ui.notify('Ya hay un procesamiento en curso', type='warning')
                            return
                        if not _state['csv_path']:
                            ui.notify('Primero sube un archivo CSV', type='warning')
                            return

                        checkpoint = logica.cargar_checkpoint()
                        usar_cp = False

                        with ui.dialog() as dlg, ui.card().classes('q-pa-md').style('min-width:380px'):
                            ui.label('¿Iniciar procesamiento?').style('font-size:1.1rem; font-weight:700')
                            ui.label(
                                f'Rango: filas {inp_fi.value} → {inp_ff.value}'
                            ).classes('text-grey-5').style('font-size:0.85rem; margin-top:4px')

                            cp_switch = None
                            if checkpoint:
                                ui.separator().classes('q-my-sm')
                                with ui.row().classes('items-center gap-2'):
                                    ui.icon('restore', color='amber').style('font-size:1.2rem')
                                    ui.label('Checkpoint encontrado').style('font-weight:600; color:#fbbf24')
                                ui.label(
                                    f"Exitosos: {checkpoint['exitosos']} · Fallidos: {checkpoint['fallidos']} · "
                                    f"Columna: {checkpoint.get('columna_actual', 'N/A')}"
                                ).classes('text-grey-5').style('font-size:0.8rem')
                                cp_switch = ui.switch('Continuar desde checkpoint', value=True).classes('q-mt-xs')

                            ui.separator().classes('q-my-sm')
                            with ui.row().classes('justify-end gap-2 q-mt-xs'):
                                ui.button('Cancelar', on_click=dlg.close).props('flat')

                                async def confirmar():
                                    nonlocal usar_cp
                                    if cp_switch is not None:
                                        usar_cp = cp_switch.value
                                    dlg.close()
                                    await _run_processing(usar_cp)

                                ui.button('▶ Procesar', on_click=confirmar).props('unelevated color=green').classes('text-white')

                        dlg.open()

                    async def _run_processing(usar_checkpoint: bool):
                        logica.procesando = True
                        logica.pausado    = False
                        logica.cancelar   = False

                        btn_procesar.set_enabled(False)
                        btn_pausar.set_enabled(True)
                        btn_cancelar.set_enabled(True)

                        log_box.clear()
                        status_badge.set_text('Procesando...')
                        status_badge.props('color=indigo outline')
                        stat_ex.set_text('0')
                        stat_fa.set_text('0')
                        stat_pr.set_text('0%')
                        progress_bar.set_value(0)
                        progress_txt.set_text('0 / 0')
                        progress_pct.set_text('0%')

                        loop = asyncio.get_event_loop()

                        def on_log(msg: str):
                            loop.call_soon_threadsafe(log_box.push, msg)

                        def on_progreso(actual: int, total: int, columna: str):
                            pct = actual / total if total > 0 else 0
                            loop.call_soon_threadsafe(progress_bar.set_value, pct)
                            loop.call_soon_threadsafe(progress_txt.set_text, f'{actual} / {total}')
                            loop.call_soon_threadsafe(progress_pct.set_text, f'{pct*100:.1f}%')
                            nombre_col = columna.upper().replace('_', ' ')
                            loop.call_soon_threadsafe(status_badge.set_text, f'Procesando {nombre_col}...')

                        logica.cb_log      = on_log
                        logica.cb_progreso = on_progreso

                        ok, mensaje, resumen = await run.io_bound(
                            logica.procesar_core,
                            _state['csv_path'],
                            inp_url.value,
                            inp_hoja.value,
                            inp_fi.value,
                            inp_ff.value,
                            usar_checkpoint
                        )

                        logica.procesando = False
                        btn_procesar.set_enabled(True)
                        btn_pausar.set_enabled(False)
                        btn_cancelar.set_enabled(False)

                        if ok and resumen:
                            stat_ex.set_text(str(resumen['exitosos']))
                            stat_fa.set_text(str(resumen['fallidos']))
                            stat_pr.set_text(f"{resumen['precision']:.1f}%")
                            progress_bar.set_value(1.0)
                            status_badge.set_text('Completado')
                            status_badge.props('color=positive outline')

                            # Resultados
                            with ui.dialog() as dlg_res, ui.card().classes('q-pa-lg').style('min-width:420px'):
                                with ui.row().classes('items-center gap-3 q-mb-sm'):
                                    ui.icon('check_circle', color='green').style('font-size:2rem')
                                    ui.label('PROCESAMIENTO COMPLETADO').style('font-size:1.2rem; font-weight:700')

                                ui.separator()

                                with ui.grid(columns=3).classes('w-full q-my-md gap-2'):
                                    for label, val, color in [
                                        ('Total', resumen['titulos_total'], '#94a3b8'),
                                        ('Exitosos', resumen['exitosos'],  '#4ade80'),
                                        ('Fallidos', resumen['fallidos'],  '#f87171'),
                                    ]:
                                        with ui.card().classes('q-pa-sm text-center').style('border-radius:10px'):
                                            ui.label(str(val)).style(f'font-size:1.6rem; font-weight:700; color:{color}')
                                            ui.label(label).classes('text-grey-5').style('font-size:0.75rem')

                                ui.label(f"Precisión global: {resumen['precision']:.1f}%").style('font-weight:600; color:#818cf8; margin-bottom:8px')

                                if resumen.get('detalle_columnas'):
                                    ui.separator()
                                    ui.label('Resultados por columna').style('font-weight:600; margin:8px 0 4px')
                                    for col, d in resumen['detalle_columnas'].items():
                                        with ui.row().classes('items-center justify-between q-py-xs'):
                                            ui.label(col).style('font-size:0.85rem; font-weight:500')
                                            ui.label(
                                                f"{d['exitosos']}/{d['total']} · {d['precision']:.1f}%"
                                            ).classes('text-green-400').style('font-size:0.85rem')

                                ui.separator()
                                ui.button('Cerrar', on_click=dlg_res.close).props('unelevated color=green').classes('text-white q-mt-sm')

                            dlg_res.open()
                            ui.notify('✅ Procesamiento completado exitosamente', type='positive', position='top', timeout=6000)

                        else:
                            tipo = 'warning' if logica.cancelar else 'negative'
                            status_badge.set_text('Cancelado' if logica.cancelar else 'Error')
                            status_badge.props(f'color={"warning" if logica.cancelar else "negative"} outline')
                            ui.notify(mensaje, type=tipo, position='top', timeout=8000)

                    btn_procesar = ui.button('▶ Procesar Datos', on_click=iniciar_procesamiento).props(
                        'unelevated color=green no-caps'
                    ).style(BTN_STYLE)

                    # ── Pausar ────────────────────────────────────────────
                    def pausar():
                        if not logica.procesando:
                            return
                        logica.pausado = not logica.pausado
                        if logica.pausado:
                            btn_pausar.set_text('▶ Reanudar')
                            btn_pausar.props('color=positive unelevated')
                            status_badge.set_text('Pausado')
                            status_badge.props('color=warning outline')
                            ui.notify('⏸ Procesamiento pausado', type='warning', position='top-right')
                        else:
                            btn_pausar.set_text('⏸ Pausar')
                            btn_pausar.props('color=warning unelevated')
                            status_badge.set_text('Procesando...')
                            status_badge.props('color=indigo outline')
                            ui.notify('▶ Procesamiento reanudado', type='positive', position='top-right')

                    btn_pausar = ui.button('⏸ Pausar', on_click=pausar).props(
                        'unelevated color=warning no-caps'
                    ).style(BTN_STYLE)
                    btn_pausar.set_enabled(False)

                    # ── Cancelar ──────────────────────────────────────────
                    async def cancelar():
                        if not logica.procesando:
                            return
                        with ui.dialog() as dlg_c, ui.card().classes('q-pa-md').style('min-width:340px'):
                            with ui.row().classes('items-center gap-2 q-mb-sm'):
                                ui.icon('warning', color='red').style('font-size:1.5rem')
                                ui.label('¿Cancelar procesamiento?').style('font-weight:700')
                            ui.label('El progreso se guardará en checkpoint').classes('text-grey-5').style('font-size:0.85rem')
                            ui.separator().classes('q-my-sm')
                            with ui.row().classes('justify-end gap-2'):
                                ui.button('No, continuar', on_click=dlg_c.close).props('flat')

                                def confirmar_cancel():
                                    logica.cancelar   = True
                                    logica.procesando = False
                                    dlg_c.close()
                                    ui.notify('🛑 Cancelando...', type='warning', position='top-right')

                                ui.button('Sí, cancelar', on_click=confirmar_cancel).props('unelevated color=red').classes('text-white')

                        dlg_c.open()

                    btn_cancelar = ui.button('🛑 Cancelar', on_click=cancelar).props(
                        'unelevated color=red no-caps'
                    ).style(BTN_STYLE)
                    btn_cancelar.set_enabled(False)

            # ── PROGRESO ──────────────────────────────────────────────────
            with ui.card().classes('app-card w-full q-pa-lg'):
                with ui.row().classes('items-center justify-between w-full q-mb-sm'):
                    ui.label('📈 Progreso').classes('section-title').style('margin-bottom:0')
                    status_badge = ui.chip('Listo', icon='check_circle', color='grey').props('outline')

                progress_bar = ui.linear_progress(value=0, size='14px', color='indigo').classes('w-full q-mb-xs').style('border-radius:8px')

                with ui.row().classes('w-full justify-between items-center'):
                    progress_txt = ui.label('0 / 0').classes('text-grey-5').style('font-size:0.82rem')
                    progress_pct = ui.label('0%').style('font-size:0.9rem; font-weight:600')

            # ── ESTADÍSTICAS ──────────────────────────────────────────────
            with ui.row().classes('w-full gap-4 q-mb-4'):
                with ui.card().classes('app-card stat-chip flex-1 bg-green-900').style('margin-bottom:0'):
                    stat_ex = ui.label('0').style('font-size:2rem; font-weight:700; color:#4ade80')
                    ui.label('Exitosos').classes('text-grey-4').style('font-size:0.8rem')

                with ui.card().classes('app-card stat-chip flex-1 bg-red-900').style('margin-bottom:0'):
                    stat_fa = ui.label('0').style('font-size:2rem; font-weight:700; color:#f87171')
                    ui.label('Fallidos').classes('text-grey-4').style('font-size:0.8rem')

                with ui.card().classes('app-card stat-chip flex-1 bg-indigo-900').style('margin-bottom:0'):
                    stat_pr = ui.label('0%').style('font-size:2rem; font-weight:700; color:#818cf8')
                    ui.label('Precisión').classes('text-grey-4').style('font-size:0.8rem')

            # Los mensajes van al archivo de log en disco
            log_box = ui.log(max_lines=1000).style('display:none')



# ──────────────────────────────────────────────────────────────────────────────
if __name__ in {'__main__', '__mp_main__'}:
    ui.run(
        title='Vuela a la Vida',
        port=8080,
        reload=False,
        dark=True,
        favicon='logo_valv.png',
        show=True,
        storage_secret=os.environ.get('STORAGE_SECRET', 'META-VALV-2026-SECRET'),
    )
