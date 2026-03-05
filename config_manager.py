"""
GESTOR DE CONFIGURACIÓN DE RUTAS - PayPal
Maneja la persistencia de rutas en config_paypal.ini
"""

import configparser
from pathlib import Path
from typing import Optional


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
        if "RUTAS_PRINCIPALES" not in self.config:
            return None

        rutas_pdf = []
        if "RUTAS_PDF" in self.config:
            total = int(self.config["RUTAS_PDF"].get("total", "0"))
            for i in range(total):
                val = self.config["RUTAS_PDF"].get(f"ruta_{i}", "").strip()
                if val:
                    rutas_pdf.append(Path(val))

        return {
            "base_paypal": Path(
                self.config["RUTAS_PRINCIPALES"].get("base_paypal", "")
            ),
            "ruta_maestro": Path(
                self.config["RUTAS_PRINCIPALES"].get("ruta_maestro", "")
            ),
            "rutas_pdf": rutas_pdf,
        }

    def guardar_config(
        self, base_paypal: str, ruta_maestro: str, rutas_pdf: list) -> bool:
        """Guarda las rutas en el archivo .ini."""
        try:
            self.config["RUTAS_PRINCIPALES"] = {
                "base_paypal": base_paypal,
                "ruta_maestro": ruta_maestro,
            }

            pdf_section: dict = {"total": str(len(rutas_pdf))}
            for i, ruta in enumerate(rutas_pdf):
                pdf_section[f"ruta_{i}"] = ruta.strip()

            self.config["RUTAS_PDF"] = pdf_section

            with open(self.CONFIG_FILE, "w", encoding="utf-8") as f:
                self.config.write(f)

            return True

        except Exception as e:
            print(f"[ConfiguradorRutasPayPal] Error al guardar config: {e}")
            return False
