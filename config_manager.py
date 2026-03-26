"""
GESTOR DE CONFIGURACIÓN DE RUTAS - PayPal
Maneja la persistencia de rutas en config_paypal.ini
"""

import configparser
from pathlib import Path
from typing import Optional


def _normalizar_ruta(valor: str) -> str:
    """
    Convierte cualquier ruta a formato Windows (backslashes).
    Si el valor está vacío o es inválido, retorna string vacío.
    """
    if not valor or not valor.strip():
        return ""
    try:
        return str(Path(valor.strip()))
    except Exception:
        return valor.strip()


def _safe_path(valor: str) -> Optional[Path]:
    """
    Convierte un string a Path solo si el valor es no vacío y no es 'None' literal.
    Retorna None si el valor es vacío, 'none', 'null', etc.
    """
    if not valor:
        return None
    limpio = valor.strip().lower()
    if limpio in ("", "none", "null", "nan"):
        return None
    return Path(valor.strip())


class ConfiguradorRutasPayPal:
    """Manejador de configuración de rutas para el sistema PayPal"""
    CONFIG_FILE = Path("config_paypal.ini")

    def __init__(self):
        self.config = configparser.ConfigParser()

    def cargar_config(self) -> bool:
        """Carga el archivo .ini. Retorna True si existe y tiene datos válidos."""
        if not self.CONFIG_FILE.exists():
            return False
        self.config.read(self.CONFIG_FILE, encoding="utf-8")
        return (
            "RUTAS_PRINCIPALES" in self.config
            and "RUTAS_PDF" in self.config
        )

    def obtener_rutas(self) -> Optional[dict]:
        """Obtiene las rutas configuradas en el archivo .ini."""
        if 'RUTAS_PRINCIPALES' not in self.config:
            return None

        rutas_pdf = []
        if 'RUTAS_PDF' in self.config:
            total = int(self.config['RUTAS_PDF'].get('total', '0'))
            for i in range(total):
                val = self.config['RUTAS_PDF'].get(f'ruta_{i}', '').strip()
                if val:
                    rutas_pdf.append(Path(val))

        # Leer raiz_swift con la clave correcta
        raiz_swift = self.config['RUTAS_PRINCIPALES'].get('raiz_swift_latam', '').strip()

        return {
            'base_paypal':      _safe_path(self.config['RUTAS_PRINCIPALES'].get('base_paypal', '')),
            'ruta_maestro':     _safe_path(self.config['RUTAS_PRINCIPALES'].get('ruta_maestro', '')),
            'rutas_pdf':        rutas_pdf,
            'raiz_swift_latam': _safe_path(raiz_swift),
        }

    def guardar_config(
        self,
        base_paypal: str,
        ruta_maestro: str,
        rutas_pdf: list,
        raiz_swift_latam: Optional[str] = None
    ) -> bool:
        """Guarda las rutas en el archivo .ini con backslashes (formato Windows)."""
        try:
            self.config["RUTAS_PRINCIPALES"] = {
                "base_paypal":      _normalizar_ruta(base_paypal),
                "ruta_maestro":     _normalizar_ruta(ruta_maestro),
                "raiz_swift_latam": _normalizar_ruta(raiz_swift_latam) if raiz_swift_latam else "",
            }

            pdf_section: dict = {"total": str(len(rutas_pdf))}
            for i, ruta in enumerate(rutas_pdf):
                pdf_section[f"ruta_{i}"] = _normalizar_ruta(str(ruta))

            self.config["RUTAS_PDF"] = pdf_section

            with open(self.CONFIG_FILE, "w", encoding="utf-8") as f:
                self.config.write(f)

            return True

        except Exception as e:
            print(f"[ConfiguradorRutasPayPal] Error al guardar config: {e}")
            return False