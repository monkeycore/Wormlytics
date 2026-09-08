# Wormlytics

Bot de auditoría de analítica digital. Le das una URL, navega el sitio por su cuenta capturando los beacons de medición y el data layer, y devuelve un Excel con todo lo encontrado — listo para documentar una implementación o levantar un SDR.

## Qué captura

- **Adobe Analytics** — AppMeasurement: eVars, props y events.
- **Adobe Web SDK (Alloy)** — incluida la extracción de objetos XDM.
- **GA4** — Measurement Protocol.
- **Data layer / TMS** — lo que la página expone en su capa de datos.
- **Eventos** capturados durante la navegación y **eventos de ecommerce**.

## Cómo funciona

Arranca el Chrome que tengas instalado con un perfil aparte y un puerto de depuración remota, y se conecta por CDP. Eso permite leer también el cuerpo de los POST, que es donde viajan Web SDK y los hits agrupados de GA4.

Desde ahí navega solo siguiendo enlaces durante el tiempo que le indiques, acumulando cada hit que detecta.

## Instalación

```bash
pip install -r requirements.txt
playwright install chromium
```

Necesitas Chrome, Chromium o Edge instalado: Wormlytics lo localiza automáticamente en Windows, macOS y Linux.

## Uso

```bash
# Navegación autónoma (5 minutos por defecto)
python wormlytics.py https://www.example.com

# Modo manual: navegas tú primero, luego toma el control el bot
python wormlytics.py https://www.example.com --manual

# Duración concreta y nombre de salida
python wormlytics.py https://www.example.com --duration 180 --output mi_analisis.xlsx

# Limitar la navegación a una sección del sitio
python wormlytics.py https://www.example.com --scope /entradas
```

| Opción | Qué hace |
|---|---|
| `--manual`, `-m` | Navegas tú primero — útil para aceptar cookies o hacer login — y después el bot sigue solo |
| `--duration`, `-d` | Segundos de navegación autónoma (por defecto 300) |
| `--output`, `-o` | Nombre del Excel de salida (por defecto `wormlytics_<dominio>_<timestamp>.xlsx`) |
| `--scope` | Restringe la navegación a las URLs que contengan ese path |
| `--port` | Puerto CDP para Chrome (por defecto 9222) |

## El Excel de salida

Un libro con una hoja por dimensión del análisis:

- **RESUMEN** — visión general de la captura.
- **VARIABLES_AA** — variables de Adobe Analytics encontradas y sus valores.
- **DATALAYER_TMS** — lo recogido de la capa de datos.
- **EVENTOS_RT** — eventos capturados durante la navegación.
- **PAGINAS** — páginas visitadas.
- **ECOMMERCE** — eventos de ecommerce.

## Aviso

Es una herramienta de auditoría: úsala sobre sitios que te corresponda analizar. El Excel que genera contiene los datos reales de la sesión capturada, así que trátalo en consecuencia y revísalo antes de compartirlo.
