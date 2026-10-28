# FONACOT · Validador de Renovaciones — v3.3

Aplicación de escritorio para Windows que arma y valida el **lote diario de cancelaciones y reversas de renovaciones**
que llegan como incidentes de iTop a Outlook, y prepara el **ticket único** del día.

La guía completa (uso diario, cómo decide la app, cómo se construye y cómo se valida) está en
[GUIA_DE_USO.md](GUIA_DE_USO.md).

---

## Qué hace

- **Lee Outlook** (todos los buzones, incluidos *Elementos enviados*) y arma el lote del día con los incidentes de
  cancelación o reversa de renovación, aunque el asunto no diga nada (*APOYO*) o la palabra venga mal escrita.
- **Desglosa cada incidente en sus renovaciones** (1 a N): varios créditos, varios clientes, tablas de CREDERE, listas,
  flechas *padre → renovación*, etiquetas como *Crédito a cancelar*, *No. de renovación* o
  *NUMERO DE CREDITO (REN)*.
- **Lee los adjuntos** cuando el correo no trae los números: capturas de pantalla, fotos (aunque estén de lado) y PDF
  como el estado de cuenta, con el OCR que ya trae Windows.
- **Compara contra los tickets únicos enviados** para saber qué ya salió, qué salió tarde, qué no salió y qué salió con
  datos equivocados; lo que no salió en su día regresa al lote de hoy.
- **Valida en la base de datos** (Oracle, solo lectura): que el cliente tenga esa renovación y de ese día, si el correo
  trae el crédito padre, si ya está en cédula y si la renovación procede.
- **Revisa el calendario de emisión de cédulas y el cierre de mes:** lo que ya se emitió no se manda y lo que pasó el
  cierre de mes necesita autorización.
- **Lee los avisos de solución de iTop** para saber qué incidentes ya se cerraron, quién los cerró y si fue por error.
- **Da un veredicto por renovación** (*Lista*, *Revisar*, *Corregir*, *Autorización*, *Atrasada · va hoy*,
  *No se manda*, *Enviada*, *Siguiente*) con las comprobaciones que lo sostienen y qué hacer.
- **Prepara el envío:** *Copiar lote* (tabla INCIDENTE | FECHA | CLIENTE | CREDITO (HIJO)), *Crear correo Ticket Único*
  (borrador en Outlook, no lo envía) y *Exportar a Excel*.
- **Corregir dato:** si el correo trae mal el cliente o el crédito, se corrige a mano o con un clic sobre lo que dice la
  base, y la corrección viaja al lote, al correo y al Excel.

## Perfiles

- **Usuario (analista):** ve *Mis renovaciones de hoy* por su gafete, las valida y las copia con
  *Copiar mis renovaciones*.
- **Validador:** arma el lote completo del día (*Por ticket único*) o el de un compañero (*Por compañero*), hace la
  *Revisión semanal* y crea el correo del ticket único.

Se elige al abrir la app y se cambia con el botón de perfil (arriba a la derecha).

## Novedades de la v3.3

- **Ejecutable para equipos sin Python:** se reparte `ValidadorRenovaciones.zip`; el analista lo descomprime y abre
  `ValidadorRenovaciones.exe`. No hay que instalar Python ni librerías.
- **Etiquetas del crédito con aclaración entre paréntesis:** *NUMERO DE CREDITO (REN) : 57815*, *CREDITO (REN): …*,
  *No. de crédito (hijo)*, *RENOVACION (CREDITO NUEVO): …* y *CREDITO (57815)* ya se leen (antes salía
  *SIN CRÉDITO*).

---

## Instalación

### Analistas (sin Python)

1. Descomprimir `ValidadorRenovaciones.zip` en **Documentos** o en el **Escritorio** (no en *Archivos de programa*).
2. Abrir `ValidadorRenovaciones\ValidadorRenovaciones.exe`.
3. Si aparece *Windows protegió su PC*: *Más información → Ejecutar de todas formas* (el `.exe` no está firmado).

El equipo necesita **Outlook** con la sesión abierta y, para validar en la base, el **Oracle Client** de SQL Developer.
Para actualizar se borra la carpeta anterior y se descomprime la nueva; la configuración se conserva.

### Desarrollo (con Python 3.10+)

```bash
pip install -r requirements.txt
```

```bash
python validador_renovaciones_app.py
```

```bash
python -m unittest discover -s tests
```

Para construir el `.zip`: doble clic en `Construir_Ejecutable.bat` (ver la guía).

---

## Estructura

| Archivo | Qué contiene |
|---|---|
| `validador_renovaciones_app.py` | Interfaz (CustomTkinter): lote, desglose, avisos, correo y Excel |
| `outlook_renovaciones_engine.py` | Lectura de Outlook, historial de asignaciones, tickets únicos y cierres de iTop |
| `core_renovaciones.py` | Lectura de los correos: clasificación y desglose de cliente y crédito |
| `ocr_imagenes.py` | Lectura de capturas, fotos y PDF con el OCR de Windows |
| `sql_renovaciones.py` | Validación en la base de datos (Oracle, solo `SELECT`) |
| `calendario_cedulas.py` | Calendario de emisión de cédulas y cierres de mes |
| `reglas_lote.py` | Qué va en el lote, veredicto y comprobaciones de cada renovación |
| `lote_utils.py` | Marcas guardadas (validadas, descartadas, correcciones) y formato del lote |
| `registro.py` | Registro de errores en `validador.log` |
| `equipo_compartido.py` | Trabajo en equipo por carpeta compartida (apagado en esta versión) |
| `tests/` | Pruebas automáticas (no necesitan Outlook ni la base) |
| `ValidadorRenovaciones.spec`, `Construir_Ejecutable.bat` | Construcción del ejecutable |
| `Ejecutar_Validador_Renovaciones.bat` | Abre el ejecutable si existe; si no, la app con Python |

## Configuración

Se guarda en `%USERPROFILE%\.fonacot_validador_renovaciones_config.json` (casi todo se cambia desde la app):

- `perfil` (`"usuario"` / `"validador"`) y `perfil_fijo` (`true` para no dejar cambiarlo).
- `gafete_usuario` / `agente_validador`: el gafete de cada perfil.
- `tema` (`"claro"` / `"oscuro"`), `letra_grande`, `columnas_ocultas`, `vista`.
- `avanzar_al_validar`, `pedir_motivo`.
- `dias_inhabiles`, `feriados_oficiales`, `dias_atras_lote` (5 días hábiles por omisión).
- `leer_imagenes`, `sql_activo`, `sql_host`, `sql_puerto`, `sql_servicio`, `sql_usuario`, `sql_tabla`, `sql_recordar`,
  `sql_cliente_oracle`. La contraseña de la base **no** se guarda aquí (si se pide, va al Administrador de
  credenciales de Windows).
- `itop_url_template`: otro servidor de iTop (admite `{id}`, `{ref}`, `{clase}`).
- `texto_correo`: texto fijo del correo del ticket único (admite `{total}`); vacío = el aprendido del último enviado.

Las marcas del día (validadas, descartadas, correcciones) se guardan en
`%USERPROFILE%\.fonacot_validador_renovaciones_estado.json` durante 30 días.
