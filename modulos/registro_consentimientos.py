"""Asociación verificable entre filas del Excel y los PDFs de una generación."""

import hashlib
import json
from pathlib import Path


def huella_registro(registro):
    contenido = json.dumps(registro, ensure_ascii=False, sort_keys=True).encode("utf-8")
    return hashlib.sha256(contenido).hexdigest()


class RegistroConsentimientos:
    NOMBRE = "registro_consentimientos.json"

    def __init__(self, carpeta):
        self.carpeta = Path(carpeta).resolve()
        self.ruta = self.carpeta / self.NOMBRE
        self.datos = None

    def iniciar(self, registros, limpiar_anteriores=False):
        self.carpeta.mkdir(parents=True, exist_ok=True)
        self.datos = {
            "version": 1,
            "completo": False,
            "registros": [huella_registro(r) for r in registros],
            "archivos": [],
        }
        self._guardar()
        eliminados = 0
        if limpiar_anteriores:
            # El lote ya está invalidado: una limpieza fallida nunca permite enviarlo.
            archivos = [
                ruta for ruta in self.carpeta.iterdir()
                if ruta.is_file() and (
                    ruta.suffix.lower() == ".pdf"
                    or ruta.name in {"emails_para_enviar.txt", "emails_lista.csv"}
                )
            ]
            for ruta in archivos:
                if ruta.resolve().parent != self.carpeta:
                    raise ValueError(f"No se puede limpiar un archivo fuera de la carpeta de salida: {ruta.name}")
            for ruta in archivos:
                try:
                    ruta.unlink()
                except OSError as exc:
                    raise ValueError(
                        f"No se pudo borrar {ruta.name}. Cierra el archivo si está abierto "
                        "y vuelve a generar los consentimientos."
                    ) from exc
                eliminados += 1
        return eliminados

    def agregar(self, numero, registro, archivo):
        ruta = Path(archivo).resolve()
        if ruta.parent != self.carpeta or ruta.suffix.lower() != ".pdf":
            raise ValueError("El PDF debe estar en la carpeta de salida.")
        if self.datos["registros"][numero - 1] != huella_registro(registro):
            raise ValueError("El PDF no corresponde al registro esperado.")
        self.datos["archivos"].append({
            "numero": numero,
            "registro": huella_registro(registro),
            "archivo": ruta.name,
            "sha256": hashlib.sha256(ruta.read_bytes()).hexdigest(),
        })
        self._guardar()

    def finalizar(self):
        numeros = [a["numero"] for a in self.datos["archivos"]]
        self.datos["completo"] = (
            bool(numeros) and sorted(numeros) == list(range(1, len(self.datos["registros"]) + 1))
        )
        self._guardar()

    def _guardar(self):
        temporal = self.ruta.with_suffix(".tmp")
        temporal.write_text(json.dumps(self.datos, ensure_ascii=False, indent=2), encoding="utf-8")
        temporal.replace(self.ruta)

    def validar(self, registros):
        """Valida el lote completo antes de permitir cualquier envío o borrador."""
        try:
            datos = json.loads(self.ruta.read_text(encoding="utf-8"))
            if datos["version"] != 1 or datos["completo"] is not True:
                raise ValueError("La generación está incompleta.")
            if not registros or datos["registros"] != [huella_registro(r) for r in registros]:
                raise ValueError("El Excel no coincide con el usado para generar los PDFs.")
            entradas = datos["archivos"]
            if len(entradas) != len(registros):
                raise ValueError("Faltan asociaciones entre registros y PDFs.")
            archivos = []
            rutas_usadas = set()
            contenidos_usados = set()
            for numero, (registro, entrada) in enumerate(zip(registros, entradas), 1):
                if entrada["numero"] != numero or entrada["registro"] != huella_registro(registro):
                    raise ValueError(f"Asociación incorrecta en el registro {numero}.")
                ruta = (self.carpeta / entrada["archivo"]).resolve()
                if ruta.parent != self.carpeta or ruta.suffix.lower() != ".pdf":
                    raise ValueError(f"Ruta de PDF incorrecta en el registro {numero}.")
                contenido = ruta.read_bytes()
                huella = hashlib.sha256(contenido).hexdigest()
                if not contenido or huella != entrada["sha256"]:
                    raise ValueError(f"El PDF del registro {numero} ha cambiado o está vacío.")
                if ruta in rutas_usadas or huella in contenidos_usados:
                    raise ValueError(f"PDF repetido en el registro {numero}.")
                rutas_usadas.add(ruta)
                contenidos_usados.add(huella)
                archivos.append({"archivo": str(ruta), "contenido": contenido})
            return archivos
        except (OSError, ValueError, KeyError, TypeError) as exc:
            raise ValueError(
                "Borradores bloqueados: no se puede verificar un PDF propio para cada registro. "
                "Vuelve a generar los consentimientos con esta versión y revisa los borradores. "
                f"Detalle: {exc}"
            ) from exc
