# Instagram Scraper + Análisis Big Five (Playwright y Gemini)

Proyecto académico en **Python** con dos etapas independientes:

1. **Extracción:** con **Playwright** recolecta los datos públicos de un perfil de Instagram: metadatos, publicaciones en alta resolución (incluidos los carruseles completos) y, opcionalmente, las listas de seguidores y seguidos.
2. **Análisis:** envía lo recolectado a la **API de Gemini** para obtener una estimación exploratoria de personalidad según el modelo **Big Five (OCEAN)**, con informes en JSON, TXT y Excel.

> Proyecto desarrollado con fines estrictamente académicos para la asignatura de Arquitectura de Software (séptimo semestre). Ver [Privacidad y uso responsable](#-privacidad-y-uso-responsable).

---

## Estructura del proyecto

```text
.
├── scraper_instagram.py          # Etapa 1: extracción con Playwright
├── analisis_bigfive_gemini.py    # Etapa 2: análisis Big Five con Gemini
├── requirements.txt              # Dependencias
├── .env.example                  # Plantilla de configuración (la clave real va en .env)
├── .gitignore                    # Excluye secretos, sesión y datos extraídos
└── datos/                        # Se genera al ejecutar (no se sube a Git)
    └── <usuario>/
        ├── perfil.json           # Metadatos: biografía, contadores
        ├── publicaciones.json    # Descripción e imágenes de cada post
        ├── enlaces.json          # Caché de URLs de publicaciones
        ├── foto_perfil.jpg
        ├── seguidores.json       # Opcional (--listas)
        ├── seguidos.json         # Opcional (--listas)
        ├── imagenes/             # post_<n>_<imagen>.jpg
        ├── bigfive.json          # Resultado estructurado + metadatos
        ├── bigfive.txt           # Informe legible
        └── bigfive.xlsx          # Tabla, gráfico y evidencias
```

---

## Requisitos e instalación

- Python **3.9 o superior** (se recomienda un entorno virtual).
- Google Chrome instalado (opcional: si no está, se usa el Chromium de Playwright).
- Una clave de la API de Gemini, generada en [Google AI Studio](https://aistudio.google.com/).

```bash
# 1. Crear y activar el entorno virtual
python -m venv venv
venv\Scripts\activate            # Windows
source venv/bin/activate         # Linux / macOS

# 2. Instalar dependencias y el navegador de Playwright
pip install -r requirements.txt
playwright install chromium
```

### Configuración de la clave de Gemini

```bash
# Copia la plantilla y pega tu clave real dentro del archivo .env
copy .env.example .env           # Windows
cp .env.example .env             # Linux / macOS
```

El archivo `.env` debe contener una sola línea:

```text
GEMINI_API_KEY=tu_clave
```

El script lo lee automáticamente (desde la carpeta de ejecución o la del script), por consola o desde el IDE. También funciona definiendo la variable de entorno `GEMINI_API_KEY`. **Nunca subas `.env` al repositorio.**

---

## Uso

### Etapa 1: extracción (`scraper_instagram.py`)

La primera vez se abre el navegador para que **inicies sesión manualmente**. Al pulsar ENTER en la consola se guardan las cookies en `cookies_sesion.json` y se reutilizan en las siguientes ejecuciones.

```bash
python scraper_instagram.py nombre_usuario                      # 25 publicaciones (por defecto)
python scraper_instagram.py nombre_usuario --cantidad 30        # cantidad personalizada
python scraper_instagram.py nombre_usuario --listas --max-lista 50   # + seguidores y seguidos
python scraper_instagram.py                                     # modo interactivo
```

| Parámetro | Descripción |
|---|---|
| `usuario` | Perfil a analizar (con o sin `@`). Si se omite, lo pregunta. |
| `--cantidad N` | Número de publicaciones (por defecto 25). |
| `--listas` | Guarda también las listas de seguidores y seguidos. |
| `--max-lista N` | Máximo de cuentas por lista (por defecto 100). |

### Etapa 2: análisis Big Five (`analisis_bigfive_gemini.py`)

Requiere haber ejecutado antes la etapa 1 sobre la misma cuenta.

```bash
python analisis_bigfive_gemini.py nombre_usuario --dry-run      # muestra qué se enviaría (no llama a la API)
python analisis_bigfive_gemini.py nombre_usuario                # análisis solo con texto
python analisis_bigfive_gemini.py nombre_usuario --con-imagenes 8   # texto + hasta 8 imágenes
```

| Parámetro | Descripción |
|---|---|
| `--dry-run` | Imprime los datos que se enviarían a Google y termina. |
| `--con-imagenes N` | Envía también hasta N imágenes (máximo 10); por defecto solo texto. |
| `--modelo NOMBRE` | Modelo de Gemini a usar. Verifica el nombre vigente en la [documentación](https://ai.google.dev/gemini-api/docs/models). |

**Qué se envía a Google:** usuario, biografía, contadores, texto de las publicaciones y, solo si se pide, algunas imágenes. No se envían enlaces ni las listas de seguidores/seguidos.

---

## Cómo está construido

- **Etapas desacopladas.** La extracción y el análisis son programas independientes que se comunican solo mediante los archivos JSON de `datos/<usuario>/`. Se puede repetir el análisis sin volver a extraer, o cambiar de proveedor de IA tocando una sola función.
- **Sesión persistente.** `storage_state` de Playwright guarda las cookies tras un login manual (no se almacenan contraseñas). El archivo de sesión equivale a una contraseña, por eso está en `.gitignore`.
- **Resistencia a cambios de la interfaz.** Las listas se abren con varias estrategias alternativas (clic en el contador, enlace, URL directa) y se leen por dos vías: el HTML del cuadro y las respuestas de red. Si todo falla, el script guarda una captura y datos de diagnóstico.
- **Imágenes completas.** Se lee el `srcset` de la imagen visible para obtener la mayor resolución y se recorren los carruseles, en lugar de usar la miniatura recortada de `og:image`.
- **Tolerancia a fallos.** Guardado incremental tras cada publicación, caché de enlaces y detención controlada si Instagram pide iniciar sesión de nuevo.
- **Salida de IA validada.** La respuesta de Gemini se pide en JSON y se valida (cinco rasgos, puntajes 0–100, niveles y confianza permitidos); si el formato falla, se reintenta.
- **Carga controlada.** Pausas aleatorias entre acciones y pausas más largas cada pocas publicaciones, para reducir el riesgo de limitación de tasa.

---

## Privacidad y uso responsable

- Úsalo **solo con cuentas propias** o de personas que hayan dado su consentimiento. Perfilar la personalidad de alguien y enviar sus publicaciones a un tercero requiere ese consentimiento.
- Las listas de seguidores y seguidos contienen datos de terceros. Por eso `datos/` no se sube al repositorio.
- `.env` y `cookies_sesion.json` equivalen a contraseñas: no los compartas ni los subas a Git. Si una clave llegara a exponerse, revócala y genera otra.
- Instagram prohíbe la recolección automatizada en sus términos de uso. Este proyecto se justifica únicamente como ejercicio académico, con pocas visitas y sobre cuentas propias.

## Limitaciones

- El análisis Big Five es una **estimación exploratoria a partir de contenido público**: refleja cómo alguien se presenta en redes, no necesariamente cómo es. No es un test psicométrico ni un diagnóstico. Para evaluarlo, se puede comparar con un cuestionario Big Five corto respondido por la propia persona.
- Los selectores de Instagram cambian sin aviso; puede ser necesario ajustarlos.
- Con pocas publicaciones con texto, la confianza de la estimación es baja.
