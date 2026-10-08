# Licitaciones de obras · Alicante y municipios limítrofes

Herramienta Python que lee hasta **10 páginas** del feed ATOM de la Plataforma
de Contratación del Sector Público siguiendo `rel="next"`. Selecciona entradas
con estado `PUB`, al menos un CPV con prefijo `45` y lugar de ejecución en
Alicante o en municipios de otras provincias que lindan con Alicante.

Genera CSV, resumen JSON y, opcionalmente, Excel con detalle, resumen de
importes, cinco mayores presupuestos y fuentes. No necesita claves ni cuenta.

## Instalación

Requiere Python 3.10 o posterior.

```bash
python -m venv .venv
```

Activa el entorno en Windows (PowerShell):

```powershell
.venv\Scripts\Activate.ps1
```

En macOS/Linux:

```bash
source .venv/bin/activate
```

Para generar Excel, instala la dependencia:

```bash
python -m pip install -r requirements.txt
```

La extracción a CSV y JSON usa solo la biblioteca estándar.

## Uso

```bash
python extraer_licitaciones.py --excel
```

Archivos en `resultados/`: `licitaciones_obras.csv`,
`licitaciones_obras.resumen.json` y `licitaciones_obras.xlsx`.
En pantalla se muestran las páginas procesadas, entradas leídas, número
de licitaciones, importe total sin IVA y las cinco mayores.

```bash
# Solo CSV y JSON, sin dependencias externas
python extraer_licitaciones.py

# Menos páginas (entre 1 y 10)
python extraer_licitaciones.py --max-paginas 3 --excel

# Solo provincia de Alicante
python extraer_licitaciones.py --adyacentes "" --excel

# Personalizar municipios adicionales
python extraer_licitaciones.py --adyacentes "Yecla;Caudete;Oliva" --excel

# Reprocesar sin red las páginas de una descarga previa
python extraer_licitaciones.py --cache --excel

# Elegir destinos de salida y caché
python extraer_licitaciones.py --salida resultados/obras.csv --directorio-cache feed_paginas --excel
```

Los archivos de salida existentes se sobrescriben. El modo `--cache` necesita
las páginas guardadas de una descarga previa; utiliza el mismo número de
páginas o uno inferior. Cada descarga admite tres intentos y un tiempo de
espera de 120 segundos por petición. Un fallo aborta la ejecución: no se
publica un resultado parcial como si fuera una descarga completa.

## Datos y criterios

CSV: expediente, órgano de contratación, objeto, CPV, presupuesto base sin
IVA, procedimiento, fecha límite de presentación, lugar de ejecución y URL.
Los campos ausentes quedan vacíos. CSV en UTF-8 con BOM, importes con punto
decimal; Excel almacena importes numéricos, fechas y enlaces clicables.

- Fuente: [feed ATOM oficial](https://contrataciondelestado.es/sindicacion/sindicacion_643/licitacionesPerfilesContratanteCompleto3.atom).
- Presupuesto: `BudgetAmount/TaxExclusiveAmount`, sin sustituirlo por valor estimado.
- CPV: proyecto o lotes. No se suman por separado presupuestos de lotes.
- Lugar: `RealizedLocation`, nunca la dirección del órgano de contratación.
- Alicante: NUTS `ES521`, nombre Alicante/Alacant o código postal `03xxx`.
- Municipios limítrofes: lista configurable `ADYACENTES` del script, con variantes de nombre. Una provincia genérica vecina no basta para seleccionar una entrada.
- Duplicados: primera aparición por ID ATOM. Entradas sin ID se conservan de forma independiente.
- `PUB` es el estado del feed: no se comprueba que la fecha límite continúe vigente ni se consulta el detalle de cada anuncio.
- Procedimiento: se conserva su código original, sin traducción.
- El resultado comprende las páginas leídas, no todas las licitaciones de España.

Comprobaciones de límites: [Registro de entidades locales de la Generalitat](https://www.entidadeslocales.gva.es/),
[memoria municipal de Muro](https://www.vilademuro.net/wp-content/uploads/2020/01/MEMORIA.pdf) y
[PGOU de Pilar de la Horadada](https://www.pilardelahoradada.org/sites/default/files/pgou/2022/textos/01_memoria_informativa.pdf).

La extracción realizada el **08/10/2026** leyó 4.887 entradas en 10 páginas:
14 licitaciones y 4.443.524,53 EUR sin IVA. Es un resultado histórico,
no una expectativa fija para ejecuciones futuras. El expediente
`2026/AR44U/00003772E` declara ciudad Murcia y CP 30201: se incluyó por la
ciudad indicada, con una nota de revisión en el Excel.

## Comprobaciones

```bash
python -m unittest discover -s tests -v
```

Las pruebas usan entradas sintéticas y no acceden a la red.

## Subir a GitHub

La carpeta local ya está inicializada como repositorio Git. El ZIP contiene
solo los archivos del proyecto. Crea un repositorio vacío en GitHub y ejecuta
dentro de esta carpeta (el primer comando también inicializa la copia del ZIP):

```bash
git init -b main
git add .
git commit -m "Proyecto inicial: extractor de licitaciones PCSP"
git remote add origin https://github.com/TU_USUARIO/TU_REPOSITORIO.git
git push -u origin main
```

Sustituye usuario y repositorio por los tuyos. Si Git solicita tu identidad,
configura `git config user.name` y `git config user.email` con tus datos.
Los resultados y la caché están excluidos mediante `.gitignore`.
