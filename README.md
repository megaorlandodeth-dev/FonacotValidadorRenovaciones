# FONACOT - Validador y Concentrador Diario de Cancelación de Renovaciones

Herramienta de escritorio en Python (`CustomTkinter` + `pywin32`) para la validación, descarte y desglose diario de incidentes de **Cancelación y Reversa de Renovaciones** asignados a través de iTop vía Outlook en FONACOT.

---

## 🎯 Objetivo

Evitar que cualquier solicitud de cancelación de renovación autorizada en el día se quede sin cancelar antes del cierre diario o cierre de mes (por ejemplo, casos con múltiples renovaciones por ticket o múltiples clientes en un solo correo).

---

## 🚀 Características Principales

1. **Desglose Inteligente 1 a N:**
   - **Tablas CREDERE / Sucursales:** Desglosa cada fila de la tabla como una renovación individual.
   - **Múltiples Créditos por Cliente:** Separa casos como `CLIENTE 5864479 CREDITO 56572 Y 56574` en filas independientes (`1 de 2`, `2 de 2`).
   - **Múltiples Clientes en un mismo correo:** Asocia cada cliente a su crédito correspondiente.
   - **Separación Crédito Padre vs Crédito Hijo:** Identifica claramente el crédito de origen anterior y el nuevo crédito/pagaré a cancelar.

2. **Dos Modos de Operación:**
   - **Modo Concentrado:** Para quien genera el lote final diario. Consolida todos los tickets del día sin duplicar la asignación inicial de Nancy (2609) con la del agente asignado final. Detecta además duplicados de renovaciones en diferentes tickets.
   - **Modo Por compañero (gafete o nombre):** Permite filtrar por analista específico respetando la regla de Nancy (2609) y mostrando alertas si el ticket fue reasignado entre compañeros.

3. **Herramientas de Cierre Diario:**
   - **Validación por fila:** marca renovaciones como validadas con doble clic, barra espaciadora o el botón *Validar*.
   - **Copiar Lote para Correo:** Copia la tabla concentrada al portapapeles (`Ctrl+V`) lista para enviar.
   - **Exportar a Excel:** Genera reporte multi-pestaña con enlaces directos a iTop y semáforo de colores.
   - **Acceso Directo a iTop:** *Abrir en iTop* abre el detalle del ticket (los `I-xxxxxx` son clase `Incident`); *Buscar el folio en iTop* (clic derecho) usa la búsqueda global por folio. Doble clic en la pestaña *Resumen por Ticket* también lo abre.

4. **Abreviaturas de una letra:** Reconoce `C5864479`, `C-56572`, `T 4530500`, `R 91258`, `Cte. 3386153`, `Cred. 61116`, `CANC REN`, etc., tanto en texto como en celdas de tablas. La letra `C` se decide por longitud: 7+ dígitos = Cliente, menos = Crédito.

5. **Corte diario y rezagadas del día anterior:**
   - El lote se envía hasta la **hora de corte** (por defecto `17:50`). Lo que llega después se marca *Después del corte* y queda fuera del lote de hoy.
   - Al extraer, el sistema escanea también el **día hábil anterior** (el lunes incluye viernes y fin de semana) y busca en todos los buzones, incluidos *Elementos enviados*, el correo cuyo asunto contiene **"ticket unico"** (configurable, sin importar acentos ni mayúsculas). Lee el cuerpo y los adjuntos Excel/CSV.
   - Toda renovación del periodo anterior cuyo **Cliente + Crédito** no aparezca en ese concentrado se agrega al lote de hoy como *Rezagada*. Que solo aparezca el número de ticket no basta (un ticket con 3 renovaciones pudo enviarse con 2).
   - Si no se encuentra el concentrado, se usa la **marca local** (la última renovación del lote que se copió o exportó en este equipo) y, si tampoco existe, la hora de corte.
   - La línea informativa muestra el concentrado usado, la **última renovación considerada** y cuántas rezagadas se agregaron. Los filtros rápidos permiten ver solo rezagadas, las ya enviadas o las que llegaron después del corte. *Restaurar* mete al lote una fila excluida.
   - La hora de la fila es la de la **primera asignación** del ticket (no la de una reasignación posterior).

6. **Interfaz v2.1 (diseño corporativo):**
   - Estilo sobrio tipo Windows/Fluent: tipografía **Segoe UI**, superficies blancas, grises neutros y un solo azul institucional; sin emojis en tablas, menús ni reportes.
   - **Barra superior** con la marca, el **reloj con cuenta regresiva al corte** (verde > 1 h, ámbar < 1 h, rojo < 15 min) y el buzón de Outlook. **Aviso a 15 min del corte** con las renovaciones pendientes.
   - **Panel lateral** con *Extraer y validar* y la configuración agrupada (modo, periodo, corte, correo, actualización automática).
   - **Encabezado "Lote del dd/mm/aaaa"** con el concentrado usado y la última renovación considerada; las acciones del lote (*Copiar lote*, *Exportar a Excel*, *Crear correo Ticket Único*) a la derecha. Un aviso ámbar aparece solo si no se encontró el concentrado del día anterior.
   - **Indicadores clicables** (Tickets, Créditos a cancelar, Múltiples, Alertas, Rezagadas, Avance con barra de progreso) que filtran la tabla.
   - **Pestañas con contador** (Desglose, Por ticket, Por compañero), buscador, filtros rápidos segmentados (Todas, Pendientes, Con alertas, Múltiples, Rezagadas, Nuevas) y menú *Otras vistas* (Validadas, Ya enviadas, Después del corte, Descartadas).
   - **Detalle en dos paneles:** el desglose del ticket y el correo original con los **números de cliente y crédito resaltados**.
   - Orden natural por columna (folios, fechas y números ordenan como número), colores suaves por estado y mensaje cuando la tabla está vacía.

7. **Formato real del "Ticket Único" (v2.2):**
   - Detecta el correo *"Cancelación de renovaciones – ticket único 22-10-2026"* y lee su tabla **INCIDENTE | FECHA | CLIENTE | CREDITO (HIJO)** renglón por renglón (incluye créditos cortos como `681`). Cuenta como enviada si coincide Cliente + Crédito, o el mismo ticket con el mismo crédito.
   - Si el ticket sí viene en el concentrado pero con otros datos, la rezagada lo dice: *"ahí viene Cliente X Crédito Y"*.
   - **Crear correo** arma el borrador con ese mismo formato: asunto `Cancelación de renovaciones – ticket único dd-mm-aaaa`, el saludo *Se envía "N" solicitudes…* y el texto *FAVOR DE BLOQUEAR…* copiados del último concentrado, y la tabla con el incidente ligado a iTop. *Copiar lote* copia esa misma tabla.

8. **Cierres en iTop (v2.2):**
   - Lee los correos *"I-xxxxxx Aviso de solución del Ticket"* y marca en violeta las renovaciones cuyo ticket ya se cerró, con **estado, fecha, quién lo cerró y el motivo** (por ejemplo *"Por instrucciones superiores, las cancelaciones de renovación solo podrán realizarse el mismo día…"*).
   - Si el aviso no dice quién lo cerró, se muestra el agente que tenía asignado el ticket en ese momento.
   - Columna *Cierre en iTop*, sección en el detalle, vista *Otras vistas → Cerradas en iTop*, cuenta como alerta, aviso al crear el correo y columnas en Excel.

9. **Funciones nuevas:**
   - **Validaciones guardadas:** las validaciones, descartes y filas incluidas a mano se guardan en disco (`%USERPROFILE%\.fonacot_validador_renovaciones_estado.json`, 30 días). No se pierden al volver a escanear ni al cerrar la app.
   - **Actualización automática** cada 5/10/15/30/60 min: vuelve a escanear Outlook, conserva tus marcas y avisa las renovaciones nuevas (filtro *Nuevas*).
   - **Crear correo Ticket Único:** abre en Outlook un borrador (no lo envía) con el formato del ticket único (punto 7), destinatarios configurables, conservando tu firma. Antes avisa si hay renovaciones pendientes o tickets ya cerrados en iTop. Así el siguiente día la app encuentra ese concentrado.
   - **Copiar lote** ahora se pega como **tabla con formato** en Outlook y como celdas en Excel.
   - **Menú de clic derecho** (copiar cliente / crédito / ticket / filas, buscar en iTop, descartar, restaurar) y **atajos de teclado** (botón *Atajos de teclado*): `F5` extraer, `Ctrl+F` buscar, `Espacio` validar, `Enter` iTop, `Supr` descartar, `Ctrl+C` copiar filas, `Ctrl+Shift+C` copiar lote, `Ctrl+M` correo, `Ctrl+E` Excel.
   - Pestaña *Por compañero* con rezagadas y pendientes por agente; el Excel incluye validadas/pendientes, filtros y encabezado fijo.

---

## 📦 Estructura del Proyecto

- `validador_renovaciones_app.py`: Interfaz gráfica moderna desarrollada en `CustomTkinter`.
- `core_renovaciones.py`: Motor de parsing, clasificación regex, normalización y desglose 1 a N de renovaciones.
- `outlook_renovaciones_engine.py`: Conector MAPI con Microsoft Outlook, escaneo multi-buzón, gestión de historial de asignaciones, búsqueda del concentrado y borrador del correo del lote.
- `lote_utils.py`: Persistencia de validaciones, tabla del lote (texto/HTML) y portapapeles con formato.
- `tests/`: Pruebas automáticas del parser, del corte/rezagadas y de las utilidades del lote (no requieren Outlook).
- `Ejecutar_Validador_Renovaciones.bat`: Script de inicio rápido.

---

## 🛠️ Requisitos e Instalación

1. **Python 3.10+** (probado en Python 3.12).
2. **Dependencias** (incluye `beautifulsoup4` y `lxml`, necesarias para leer las tablas HTML de iTop):
   ```bash
   pip install -r requirements.txt
   ```
3. **Microsoft Outlook** con sesión activa en Windows.
4. **Ejecución:** doble clic en `Ejecutar_Validador_Renovaciones.bat` (instala dependencias si faltan), o:
   ```bash
   python validador_renovaciones_app.py
   ```
5. **Pruebas** (no requieren Outlook):
   ```bash
   python -m unittest discover -s tests -v
   ```

### Configuración

Se guarda en `%USERPROFILE%\.fonacot_validador_renovaciones_config.json`. Además de lo que se captura en pantalla:
- `itop_url_template`: para otro servidor de iTop, p. ej. `https://servidor/pages/UI.php?operation=details&class={clase}&id={id}` (admite `{id}`, `{ref}`, `{clase}`). Vacío = servidor de FONACOT.
- `ultimo_lote`: marca de la última renovación considerada (se actualiza al copiar/exportar el lote).
- `texto_correo`: texto fijo del correo del ticket único (admite `{total}`). Vacío = usar el aprendido.
- `texto_correo_aprendido`: saludo y avisos copiados del último concentrado encontrado (se actualiza solo).
- El asunto anterior `TICKET UNICO {fecha}` se cambia automáticamente al nuevo `Cancelación de renovaciones – ticket único {fecha}`.

