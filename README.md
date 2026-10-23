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
   - **El corte de cada día es la hora en que realmente salió su ticket único**, no una hora fija: si un día se mandó a las 16:00, lo que llegó después va al día siguiente; si se mandó a las 19:10, lo que llegó a las 18:30 sí debía ir en él.
   - Mientras el ticket único de hoy no sale, todo lo que llega entra en el lote de hoy. Cuando la app lo encuentra, lo que llega después se marca *Después del envío · va el dd/mm*, y lo de hoy que no venga en él se avisa como *falta en el de las hh:mm*.
   - Lo que llegó en los 30 minutos previos al envío y no viene en él se toma como *llegó junto al envío* (va al día siguiente, sin marcarse como NO ENVIADA).
   - Las respuestas o reenvíos del concentrado (*RE:*, *RV:*, *FW:*) no cambian la hora de envío.
   - La **hora habitual de envío** (por defecto `17:50`) solo se usa como recordatorio y para los días en que no se encuentra el ticket único.
   - Al extraer, el sistema escanea también el **día hábil anterior** (el lunes incluye viernes y fin de semana) y busca en todos los buzones, incluidos *Elementos enviados*, el correo cuyo asunto contiene **"ticket unico"** (configurable, sin importar acentos ni mayúsculas). Lee el cuerpo y los adjuntos Excel/CSV.
   - Toda renovación del periodo anterior cuyo **Cliente + Crédito** no aparezca en ese concentrado se agrega al lote de hoy como *Rezagada*. Que solo aparezca el número de ticket no basta (un ticket con 3 renovaciones pudo enviarse con 2).
   - Si no se encuentra el concentrado, se usa la **marca local** (la última renovación del lote que se copió o exportó en este equipo) y, si tampoco existe, la hora habitual.
   - La línea informativa muestra el concentrado usado, la **última renovación considerada** y cuántas rezagadas se agregaron. Los filtros rápidos permiten ver solo rezagadas, las ya enviadas o las que van en el siguiente lote. *Restaurar* mete al lote una fila excluida.
   - La hora de la fila es la de la **primera asignación** del ticket (no la de una reasignación posterior).

6. **Interfaz v2.1 (diseño corporativo):**
   - Estilo sobrio tipo Windows/Fluent: tipografía **Segoe UI**, superficies blancas, grises neutros y un solo azul institucional; sin emojis en tablas, menús ni reportes.
   - **Barra superior** con la marca, el **reloj con el estado del ticket único de hoy** (pendiente con cuenta regresiva a la hora habitual, o *enviado a las hh:mm*) y el buzón de Outlook. **Aviso a 15 min de la hora habitual** si el ticket único de hoy aún no sale.
   - **Panel lateral** con *Extraer y validar* y la configuración agrupada (modo, periodo, envío del ticket único, correo, actualización automática).
   - **Encabezado "Lote del dd/mm/aaaa"** con el concentrado usado y la última renovación considerada; las acciones del lote (*Copiar lote*, *Exportar a Excel*, *Crear correo Ticket Único*) a la derecha. Un aviso ámbar aparece solo si no se encontró el concentrado del día anterior.
   - **Indicadores clicables** (Tickets, Créditos a cancelar, Múltiples, Alertas, Rezagadas, Avance con barra de progreso) que filtran la tabla.
   - **Pestañas con contador** (Desglose, Por ticket, Por compañero), buscador, filtros rápidos segmentados (Todas, Pendientes, Con alertas, Múltiples, Rezagadas, Nuevas) y menú *Otras vistas* (Validadas, Ya enviadas, Para el siguiente lote, Descartadas).
   - **Detalle en dos paneles:** el desglose del ticket y el correo original con los **números de cliente y crédito resaltados**.
   - Orden natural por columna (folios, fechas y números ordenan como número), colores suaves por estado y mensaje cuando la tabla está vacía.

7. **Formato real del "Ticket Único" (v2.2):**
   - Detecta el correo *"Cancelación de renovaciones – ticket único 22-10-2026"* y lee su tabla **INCIDENTE | FECHA | CLIENTE | CREDITO (HIJO)** renglón por renglón (incluye créditos cortos como `681`). Cuenta como enviada solo si coincide **Cliente + Crédito** (si al incidente le falta el cliente, basta el mismo ticket con el mismo crédito). Si el ticket único trae el incidente con otro cliente o crédito, se marca como **datos equivocados** (ver punto 11).
   - Si el ticket sí viene en el concentrado pero con otros datos, la rezagada lo dice: *"ahí viene Cliente X Crédito Y"*.
   - **Crear correo** arma el borrador con ese mismo formato: asunto `Cancelación de renovaciones – ticket único dd-mm-aaaa`, el saludo *Se envía "N" solicitudes…* y el texto *FAVOR DE BLOQUEAR…* copiados del último concentrado, y la tabla con el incidente ligado a iTop. *Copiar lote* copia esa misma tabla.

8. **Cierres en iTop (v2.2):**
   - Lee los correos *"I-xxxxxx Aviso de solución del Ticket"* y marca en violeta las renovaciones cuyo ticket ya se cerró, con **estado, fecha, quién lo cerró y el motivo** (por ejemplo *"Por instrucciones superiores, las cancelaciones de renovación solo podrán realizarse el mismo día…"*).
   - Si el aviso no dice quién lo cerró, se muestra el agente que tenía asignado el ticket en ese momento.
   - Columna *Cierre en iTop*, sección en el detalle, vista *Otras vistas → Cerradas en iTop*, cuenta como alerta, aviso al crear el correo y columnas en Excel.

9. **Pantalla configurable (v2.3):**
   - Botón **Vista** (o clic derecho en el encabezado de la tabla): elige qué columnas se ven y oculta los indicadores, el detalle del ticket o la configuración lateral para dar más espacio a la tabla. Se recuerda entre sesiones; *Restablecer vista* regresa a lo predeterminado.
   - De inicio se ocultan *Crédito origen*, *Asignación* y *Motivo* (siguen en el detalle y en el Excel).
   - Textos cortos en la tabla (*Rezagada 22/10 14:53 · otros datos en el concentrado*, *Resuelto 22/10 17:52 · 7353 Sammy*); al pasar el mouse sobre la celda se ve el texto completo.
   - Reconoce el formato **Cliente\*Crédito** (*"Cancelación de crédito 1184282\*10430"*).

10. **Revisión de envíos de varios días (v2.3):**
   - Cada renovación se compara contra **todos** los ticket únicos del periodo, no solo el de ayer. Su día de lote es el día en que llegó (o el siguiente día hábil si llegó después de que salió el ticket único de ese día, o en fin de semana).
   - Si salió en el ticket único de su día, o en uno posterior, ya **no cuenta como faltante** (aunque el incidente se cierre después). Ejemplo: llegó el lunes, no se mandó el lunes pero sí el miércoles → *Enviada tarde 07/10 · debía ir 05/10*.
   - Si nunca se envió:
     - con el incidente **abierto** → *NO ENVIADA (30/09) · va hoy*, se suma al lote de hoy;
     - con el incidente **cerrado** en iTop → *NO ENVIADA (30/09) · cerrado el 01/10*, en rojo y fuera del lote (*Restaurar* la mete);
     - si de ese día no se encontró ticket único → *Sin comprobar*.
   - Para revisar varios días usa **Revisión semanal** (punto 11) o elige el periodo (*Del / Al*). El título cambia a *Revisión de envíos del dd/mm al dd/mm/aaaa* con el resumen (enviadas, tarde, no enviadas, sin comprobar).
   - **Aviso rojo** clicable con las no enviadas, y vistas *Otras vistas → No enviadas en su día / Enviadas tarde / Sin comprobar*. El detalle del ticket dice en rojo en qué ticket único debía ir.
   - El Excel agrega la hoja **5_Revision_de_Envios** (resultado, día que debía ir, en qué ticket único salió, días de retraso y cierre en iTop).

11. **Lote del día o revisión semanal (v2.4):**
   - Arriba del panel lateral, **¿Qué quieres revisar?**: *Lote del día* o *Revisión semanal*. Se elige el modo, ya no depende de las fechas.
   - **Lote del día para cualquier fecha:** con *Lote del día* elige el día en *Del* o *Al* (las dos fechas se mueven juntas). Si es un día pasado, la app arma el lote **como debía salir ese día**: lo compara solo contra el ticket único de ese día, marca *falta en el de las hh:mm* lo que no venía y anota *salió el dd/mm* si se mandó después. *Hoy* regresa al lote de hoy.
   - **Revisión semanal** (*Esta semana* / *Semana pasada*): revisa de lunes a viernes, abre directo la vista *No enviadas en su día* y cuenta solo lo que debía salir esa semana (lo de la semana anterior que se escaneó para las rezagadas no se suma).
   - **Pestaña «Tickets únicos»:** junta en una sola lista todo lo que salió en los tickets únicos del periodo (y los posteriores, donde llegan las que se mandaron tarde) y revisa cada renglón contra los incidentes de Outlook: *Bien*, **DATOS EQUIVOCADOS**, *no se encontró el correo del incidente*, *repetida en N tickets únicos*. Doble clic abre el incidente en iTop.
   - **Datos equivocados en el ticket único:** si el ticket único trae un incidente con el cliente de otro incidente (o con otro crédito), esa renovación **no cuenta como enviada**: sale en rojo *NO ENVIADA · datos equivocados* con el detalle (*salió con Cliente 5864479 en lugar de 4530500 (ese cliente es del incidente I-361001)*). Si Cliente + Crédito están bien pero con el folio de otro incidente, cuenta como enviada pero se marca el error. Vista *Otras vistas → Datos equivocados en el ticket único* y aviso rojo.
   - El Excel agrega la hoja **6_Enviado_en_Tickets_Unicos** y la columna *Datos equivocados en el ticket único* en la hoja 5.
   - **Pestaña «Por día»** (en la revisión): un renglón por día hábil con el ticket único que salió (hora, quién, renglones) y cuántas debían ir, a tiempo, tarde, NO enviadas, datos equivocados, sin comprobar y pendientes. Rojo = faltó algo, ámbar = día sin ticket único encontrado. Doble clic en un día abre su *Lote del día* tal como debía salir. El encabezado avisa los *Días sin ticket único encontrado*.

**Mejoras de la auditoría (v2.4):**
   - **Días festivos:** los feriados oficiales (1 ene, 1er lunes de feb, 3er lunes de mar, 1 may, 16 sep, 3er lunes de nov, 25 dic) y los que captures en *Días inhábiles* (`02/11/2026, 12/12/2026`) no cuentan como día hábil: lo que llega en festivo va al siguiente día hábil. Si ese día sí salió un ticket único, se toma como día trabajado.
   - **Sin comprobar en vez de dar por enviada:** si no se encontró el ticket único de un día, sus renovaciones ya no se dan por enviadas; salen como *Sin comprobar* para que las revises.
   - **Fecha del asunto corregida:** si el asunto trae una fecha futura o la de un día que ya tenía su ticket único (p. ej. el del 22 dice "21-10"), se toma el día en que se envió y se avisa *Fecha corregida*.
   - **Lo no enviado no se pierde:** el lote del día revisa 5 días hábiles hacia atrás (`dias_atras_lote`), así una renovación que no salió sigue apareciendo hasta que se envíe.
   - **Rezagadas que tampoco vinieron:** si el ticket único del día ya salió y no trae una rezagada, también se marca *falta en el de las hh:mm* y cuenta en *faltaron*.
   - **Adjuntos:** si la relación del ticket único viene en un Excel/CSV adjunto, se lee renglón por renglón igual que la tabla del correo.
   - **Copias del mismo correo** (Enviados y Bandeja de varios compañeros, con uno o dos minutos de diferencia) cuentan una sola vez.
   - **Solicitudes duplicadas:** si una renovación ya había salido en otro ticket único antes de que llegara el incidente, se avisa *posible duplicado*; dos incidentes que piden lo mismo no se marcan como datos equivocados.
   - **El rojo gana:** una renovación que no se envió o salió mal se ve en rojo aunque la hayas validado antes.
   - La marca *última considerada* solo la deja el lote de hoy al copiarlo o crear el correo (no el Excel, ni una revisión, ni un lote pasado).

12. **Funciones nuevas:**
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
- `ultimo_lote`: marca de la última renovación considerada (se actualiza al copiar el lote o crear el correo de hoy).
- `dias_inhabiles` / `feriados_oficiales`: días inhábiles capturados en pantalla y si se suman los feriados oficiales.
- `dias_atras_lote`: días hábiles hacia atrás que revisa el lote del día (5 por omisión).
- `columnas_ocultas` y `vista`: columnas y zonas de la pantalla elegidas en *Vista*.
- `texto_correo`: texto fijo del correo del ticket único (admite `{total}`). Vacío = usar el aprendido.
- `texto_correo_aprendido`: saludo y avisos copiados del último concentrado encontrado (se actualiza solo).
- El asunto anterior `TICKET UNICO {fecha}` se cambia automáticamente al nuevo `Cancelación de renovaciones – ticket único {fecha}`.

