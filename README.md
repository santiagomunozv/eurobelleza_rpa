# eurobelleza_rpa

Proceso batch para Windows que:

1. Lee archivos `.PE0` nuevos desde `s3://eurobelleza-siesa/pedidos/`
2. Los deja en la ruta de Siesa
3. Ejecuta la importación en Siesa 8.5 por teclado
4. Sube los `.P99` generados a `errores/`
5. Sube un JSON de resultado a `resultados/`

El bot ya no corre en loop infinito. La idea es ejecutarlo una vez por corrida desde el Programador de tareas de Windows.

## Requisitos

- Windows con acceso a Siesa y a la unidad `U:`
- Python 3.11+ recomendado
- Acceso del usuario Windows a S3
- Sesión disponible y sin interferencia humana durante la corrida

## Instalación

```powershell
python -m pip install -r requirements.txt
```

## Configuración

Editar [config.py](./config.py) y ajustar:

- rutas reales del acceso directo y carpetas de Siesa
- usuario y contraseña de Siesa
- credenciales AWS
- título de ventana de Siesa si cambia
- secuencias de teclado si el menú cambia

Puntos importantes:

- `DELETE_SOURCE_OBJECTS` queda en `False` por defecto porque la policy actual de Windows no tiene `DeleteObject` sobre `pedidos/`.
- El bot guarda un `state.json` local para no reprocesar los mismos objetos de S3 en cada corte.

## Ejecución manual

```powershell
python bot.py
```

Cada ejecución:

- adquiere un lock local para evitar doble corrida
- genera un log en `C:\eurobelleza_rpa\logs`
- genera un JSON de resultado en `C:\eurobelleza_rpa\archive`
- sube el resultado a `s3://eurobelleza-siesa/resultados/`

## Programador de tareas

Crear una tarea que ejecute el bot en estos horarios:

- `6:30 AM`
- `1:00 PM`
- `7:00 PM`

Condiciones recomendadas:

- ejecutar solo con la sesión correcta
- no lanzar una nueva instancia si la anterior sigue corriendo
- usar una cuenta con acceso a `U:` y a S3

## Empaquetado

```powershell
Remove-Item -Recurse -Force build -ErrorAction SilentlyContinue
Remove-Item -Recurse -Force dist -ErrorAction SilentlyContinue
python -m PyInstaller Bot.spec
```

El ejecutable queda en `dist\Bot.exe`.

## Salida esperada a S3

### `errores/`

Archivos `.P99` renombrados por corrida.

### `resultados/`

JSON con esta forma:

```json
{
  "run_id": "20260406_190000",
  "started_at": "2026-04-06T19:00:01-05:00",
  "finished_at": "2026-04-06T19:12:33-05:00",
  "machine_name": "PC-SIESA-01",
  "files_detected": ["00003663.PE0"],
  "files_attempted": ["00003663.PE0"],
  "files_without_error": ["00003663.PE0"],
  "files_with_error": [],
  "fatal_error": null
}
```
