# Guía del Validador de Renovaciones (v3.3)

1. [Instalar y abrir](#1-instalar-y-abrir)
2. [Uso diario](#2-uso-diario)
3. [Cómo valida la app cada renovación](#3-cómo-valida-la-app-cada-renovación)
4. [Cómo se construye el ejecutable](#4-cómo-se-construye-el-ejecutable)
5. [Cómo se valida un cambio antes de repartirlo](#5-cómo-se-valida-un-cambio-antes-de-repartirlo)
6. [Problemas comunes](#6-problemas-comunes)

---

## 1. Instalar y abrir

**Analistas (sin Python):**

1. Copiar `ValidadorRenovaciones.zip` y descomprimirlo en **Documentos** o en el **Escritorio**. No en *Archivos de
   programa*: la app guarda junto al `.exe` su registro (`validador.log`) y el calendario de cédulas.
2. Abrir `ValidadorRenovaciones\ValidadorRenovaciones.exe`. Conviene anclarlo a la barra de tareas.
3. Si sale *Windows protegió su PC*: *Más información → Ejecutar de todas formas*. Si Sistemas bloquea programas fuera
   de *Archivos de programa*, hay que pedirles que lo permitan.
4. La primera vez se elige el perfil (**Usuario** o **Validador**) y el gafete.
5. Opcional, pero recomendable la primera vez: doble clic en **`Revisar_Equipo.bat`** (junto al `.exe`). Revisa
   que el equipo tenga todo y muestra el resultado; debe decir *0 falla(s)* (ver [5.3](#53-probar-el-ejecutable)).

El equipo necesita **Outlook** abierto con la cuenta donde llegan los incidentes. Para validar en la base hace falta un
**Oracle Client** de 64 bits (SQL Developer no lo trae: se instala aparte) y la VPN o la red de FONACOT. La lectura de capturas y PDF usa lo que ya trae
Windows 10/11.

**Actualizar:** cerrar la app, borrar la carpeta `ValidadorRenovaciones` y descomprimir el `.zip` nuevo en el mismo
lugar. La configuración y las marcas del día están en el perfil de Windows y no se pierden.

---

## 2. Uso diario

### Analista (perfil Usuario)

1. **Extraer y validar** (F5): aparecen *Mis renovaciones de hoy* (las asignadas a su gafete).
2. Revisar cada renovación con **Siguiente pendiente** (N). A la derecha está el desglose: el veredicto, qué hacer y
   las comprobaciones; abajo, el correo original con los números resaltados.
3. Si todo cuadra, **Espacio** la marca como validada y pasa a la siguiente.
4. Si el cliente o el crédito vienen mal o faltan: **Corregir dato** (E). Se escribe el bueno o se elige la renovación
   que propone la base.
5. Al terminar: **Copiar mis renovaciones** y pegarlas a quien manda el ticket único.

### Validador

1. Elegir **Lote del día** (o un día pasado para ver cómo debía salir) y **Extraer y validar** (F5).
2. Revisar los **avisos** de arriba: los rojos cambian el lote de hoy; los ámbar son para revisar. Un clic en un aviso
   deja en la tabla solo esas renovaciones.
3. Los **indicadores** (*En el lote*, *Listas*, *Revisar*, *Ya enviadas*, *No se mandan*, *Avance*) filtran la tabla
   con un clic.
4. Trabajar las de **Revisar** y **Corregir** hasta que todo quede *Lista* o validado. **Validar visibles** valida en
   bloque (con vista previa; deja fuera lo que hay que revisar).
5. **Crear correo Ticket Único** (Ctrl+M) abre el borrador en Outlook con el saludo y la tabla. Revisarlo y enviarlo
   desde Outlook. *Copiar lote* (Ctrl+Shift+C) copia solo la tabla.
6. Una vez por semana: **Revisión semanal** y las pestañas **Tickets únicos** y **Por día** para ver qué no salió,
   qué salió tarde y qué salió con datos equivocados.

### Atajos

| Tecla | Acción |
|---|---|
| F5 | Extraer y validar |
| Espacio / doble clic | Validar o quitar la validación (y pasar a la siguiente) |
| N | Siguiente pendiente |
| J / K | Renovación siguiente / anterior |
| E | Corregir dato (cliente o crédito) |
| Enter | Abrir el incidente en iTop |
| Supr | Descartar |
| Ctrl+Z | Deshacer |
| Ctrl+C / Ctrl+Shift+C | Copiar filas / copiar el lote |
| Ctrl+M / Ctrl+E | Crear correo del ticket único / exportar a Excel |
| Ctrl+F / Ctrl+K | Buscar / buscar una acción |
| Ctrl+B | Ocultar o mostrar el panel izquierdo |
| ? | Ver los atajos |

---

## 3. Cómo valida la app cada renovación

### 3.1 De dónde salen los datos

1. **Correo del incidente** (Outlook): asunto, descripción y tablas. Se desglosa en una fila por renovación
   (cliente + crédito a cancelar). Si el correo trae el **crédito padre**, se cambia por su renovación.
2. **Adjuntos:** si al correo le falta el cliente o el crédito, se leen sus capturas, fotos y PDF (dos primeras
   páginas). En un **estado de cuenta** la renovación es el crédito con la *FECHA AUTORIZADO* más reciente, hasta un
   mes antes del correo.
3. **Historial de asignación:** quién recibió el incidente primero (la hora de la fila es la de la primera
   asignación) y si se reasignó.
4. **Tickets únicos enviados** (cualquier buzón, también *Elementos enviados*), con su tabla o su Excel adjunto.
5. **Avisos de solución de iTop:** si el incidente se cerró, cuándo, quién y con qué motivo.
6. **Base de datos** (si está activa): `TOPAZ.CR_RENOVACIONES` y, para renovaciones de otro día, `SALDOS`,
   `CR_DISPOSICIONES` y `CR_CEDULAS_DETALLE`. Solo consultas `SELECT` con variables.
7. **Calendario de emisión de cédulas** (página de FONACOT, se guarda en `calendario_cedulas.json`) y cierres de mes.

### 3.2 Las comprobaciones

Cada renovación muestra estas comprobaciones con su evidencia (✔ bien, ⚠ revisar, ✖ corregir, • dato):

| Comprobación | ✔ | ⚠ | ✖ |
|---|---|---|---|
| **Correo** | trae cliente y crédito | crédito padre cambiado, leído de una imagen o PDF, dos clientes, otro número de 7 dígitos sin leer | falta el cliente o el crédito |
| **Base de datos** | la renovación existe y es de ese día | otra fecha, sugerencia, crédito padre | no existe, no es de ese cliente, no procede, ya está cancelada (`TZ_LOCK` distinto de 0) |
| **Calendario** | no ha pasado emisión ni cierre | pasó el cierre de mes (autorización), falta la fecha de la siguiente emisión | ya se emitió o ya está en la cédula |
| **Ticket único** | va en el de hoy | no salió en su día (va hoy), posible duplicado | salió con datos equivocados |
| **iTop / asignación** | — | cerrado por error, posible cierre por error, resuelto sin cancelar, reasignación, duplicado en otro incidente | — |

### 3.3 El veredicto

Cada renovación tiene **una sola decisión**. Gana la primera que aplica, en este orden:

1. **Descartada** (la descartó el validador).
2. **Corregir · salió con datos equivocados** en un ticket único.
3. **No se manda**: ya se emitió, ya está cancelada en la base (`TZ_LOCK` distinto de 0), se resolvió en iTop o la
   base dice que no procede.
4. **¿Se cerró por error?** (no se pudo leer `TZ_LOCK`) o **¿Resuelto sin cancelar?** (se resolvió en iTop, pero la
   renovación sigue sin bloquear y sí procede): Revisar. Si se debe cancelar, *Restaurar* la manda hoy.
5. **Enviada** / **Siguiente** (va en el siguiente ticket único) / **Sin comprobar** (no se encontró el ticket único
   de su día).
6. Si va en el ticket único de hoy:
   - con alguna comprobación ✖ → **Corregir** («Corrige antes de enviar»);
   - pasó el cierre de mes → **Autorización**;
   - no salió en su día o se cerró por error → **Atrasada · va hoy**;
   - con alguna ⚠ → **Revisar**;
   - todo ✔ → **Lista** (al validarla, *✓ Validada*).

### 3.4 Qué día le toca

- El corte de cada día es la **hora en que salió su ticket único**: lo que llega después va al siguiente día hábil.
- Los fines de semana y festivos (oficiales y los de *Días inhábiles*) no son días hábiles, salvo que ese día haya
  salido un ticket único (guardia).
- El lote de hoy revisa **5 días hábiles hacia atrás**: lo que no salió en su día y sigue abierto regresa como
  *Atrasada · va hoy*.
- Una renovación cuenta como enviada solo si el ticket único trae **el mismo cliente y crédito**. Si trae otros datos,
  es *datos equivocados* y no cuenta.

---

## 4. Cómo se construye el ejecutable

Solo lo hace quien reparte la app, en un equipo con **Python 3.10+** (probado con 3.11 y 3.12) e internet:

1. Bajar la rama de la versión (por ejemplo `fix/v3.3`).
2. Doble clic en **`Construir_Ejecutable.bat`**:
   - instala las librerías de `requirements.txt` y PyInstaller;
   - arma `dist\ValidadorRenovaciones\` con `ValidadorRenovaciones.spec` (Python, CustomTkinter, pandas, pywin32,
     oracledb, Pillow, lxml y los datos de idioma van adentro);
   - genera esta guía en PDF (`GUIA_DE_USO.pdf`, con `guia_pdf.py` y Microsoft Edge) y la pone junto al `.exe`;
   - lo comprime en **`ValidadorRenovaciones.zip`** (unos 56 MB).
3. Probar el `.zip` (sección 5.3) y repartirlo (carpeta compartida, correo interno o *release* en GitHub).

Notas:

- `dist\`, `build\` y el `.zip` no se suben al repositorio. El PDF de la guía sí, para descargarlo desde GitHub.
- Si se cambia la guía (`GUIA_DE_USO.md`), el PDF se vuelve a generar con `python guia_pdf.py`.
- No usar un ejecutable de un solo archivo: tarda más en abrir y los antivirus lo marcan más.
- Al descomprimir, la carpeta donde quede no debe ser muy profunda: la ruta más larga dentro del paquete mide unos
  110 caracteres y Windows tiene un límite de 260.
- Para ver la versión del `.exe`: pie de la ventana (*Versión 3.3.5*). Se cambia en `APP_VERSION` de
  `validador_renovaciones_app.py`.

---

## 5. Cómo se valida un cambio antes de repartirlo

### 5.1 Pruebas automáticas

No necesitan Outlook ni la base:

```bash
python -m unittest discover -s tests
```

Cada versión tiene su archivo (`tests/test_v33.py`, …) con los correos reales que fallaron y lo que debe salir. Un
cambio solo se reparte si **todas** pasan.

### 5.2 Un correo que se leyó mal

1. Copiar el texto del correo (sin datos de más) en una prueba nueva de `tests/test_vXX.py`, con el cliente y el crédito
   que deben salir.
2. Correr la prueba: debe fallar (así se confirma que reproduce el error).
3. Corregir la lectura en `core_renovaciones.py` (etiquetas y patrones de cliente y crédito) o en `ocr_imagenes.py`
   (capturas y PDF).
4. Correr **todas** las pruebas: la corrección no debe romper otro formato que ya se leía.

### 5.3 Probar el ejecutable

1. Descomprimir el `.zip` en una carpeta nueva (por ejemplo en el Escritorio) de un equipo **sin Python**.
2. Doble clic en **`Revisar_Equipo.bat`** (o `ValidadorRenovaciones.exe --autoprueba`). Prueba dentro del
   ejecutable, sin usuario ni contraseña: las librerías, el cifrado y el conector de la base (modo thin), el Oracle
   Client (modo clásico), el Excel, el OCR con una imagen de prueba, el lector de PDF con esta guía, Outlook, que se
   pueda escribir en la carpeta y la página del calendario. Deja el resultado en `autoprueba.txt`. Con alguna
   **FALLA** no se reparte; los **AVISO** dependen del equipo (Outlook cerrado, sin Oracle Client, sin internet).
3. Abrir el `.exe`: debe salir la ventana con la versión al pie.
4. **Extraer y validar** con Outlook abierto: deben aparecer las renovaciones del día.
5. Con *Validaciones automáticas*: **Probar conexión** a la base y revisar que la columna *Base de datos* se llene.
6. Abrir un incidente con captura o PDF y revisar que el desglose diga *leído de la imagen* o *del estado de cuenta*.
7. **Exportar a Excel** y **Crear correo Ticket Único** (solo el borrador, sin enviarlo).

Si al ejecutable le falta una librería que la app carga solo en algún momento (como `cryptography`, que la base
usa al conectarse), se agrega en `hiddenimports` de `ValidadorRenovaciones.spec` y se prueba de nuevo con la
autoprueba.

### 5.4 En el día a día

- Lo que la app no pudo leer sale **en rojo** con *Corregir* («falta el crédito»). El correo original está abajo con
  los números resaltados: con **Corregir dato** se completa y se manda una captura del correo al responsable de la
  app para que agregue ese formato.
- Lo leído de imágenes o PDF siempre queda **en ámbar** para compararlo contra el adjunto.
- Los errores quedan en `validador.log`, junto al `.exe`.

---

## 6. Problemas comunes

| Mensaje o síntoma | Qué hacer |
|---|---|
| *Windows protegió su PC* | *Más información → Ejecutar de todas formas*. |
| No aparece ninguna renovación | Revisar que Outlook esté abierto y el periodo elegido; F5 otra vez. |
| *SIN CRÉDITO* / *SIN CLIENTE* | Ver el correo original o el adjunto y usar **Corregir dato** (E). |
| *DPI-1050* al conectar a la base | El Oracle Client del equipo es 12.2: la app ya usa oracledb 3.4, compatible. Si persiste, poner la carpeta del Oracle Client en `sql_cliente_oracle`. |
| Captura de *Consulta de Crédito*: lee el crédito pero sale *SIN CLIENTE* | Pasaba antes de la 3.3.3 (el OCR se saltaba el renglón del cliente). Con la 3.3.3 se lee; si aun así falta, la columna *Base de datos* dice de qué cliente es la renovación: usar **Corregir dato** (E). |
| Correo con «Crédito / Renovación / Cliente» (el crédito primero): sale el crédito como cliente o se manda el crédito en vez de la renovación | Pasaba antes de la 3.3.4. Con la 3.3.4 se cancela la renovación y el crédito queda como padre. |
| Se mandó una renovación que ya estaba cancelada | Desde la 3.3.4, con la base activa, la que ya está bloqueada (`TZ_LOCK` distinto de 0) sale del lote: *No se manda: ya está cancelada*. |
| Al cambiar de Validador a Usuario (o al revés) seguía haciendo lo del otro perfil | Pasaba antes de la 3.3.4. Al cambiar de perfil se limpia la tabla: F5 para buscar con el perfil nuevo. |
| En una laptop no se ve el correo original (o algo queda cortado abajo) | Desde la 3.3.5 sale una barra a la derecha para bajar. El correo también se abre completo con **Abrir en ventana ↗** o con doble clic. |
| No aparece un incidente que dice «cancelar el crédito X y regresar al Y» | Pasaba antes de la 3.3.5 (no decía «renovación»). Con la 3.3.5 se cancela X y Y queda como crédito origen. |
| *ORA-12154* en modo clásico | Pasaba antes de la 3.3.3 con algunos Oracle Client instalados: descargar la versión más reciente. |
| *DPY-4000* / *getaddrinfo failed* | Revisar *Nombre del host* (`scanfonacot.fonacot.gob.mx`), *Puerto* y *Nombre del servicio*, y la VPN. |
| *no tiene la fecha de la emisión de …* | La página de FONACOT no ha publicado esa emisión o no hubo internet; se vuelve a intentar en la siguiente búsqueda. |
| *DPY-3015 … password verifier type … not supported* y *no hay un Oracle Client* | La base pide el modo clásico. SQL Developer no sirve para esto (usa su propio driver de Java): hay que instalar Oracle Instant Client 19 de 64 bits en cualquier disco, o poner su carpeta en `sql_cliente_oracle`. `Revisar_Equipo.bat` dice si lo encuentra. |
| *ORA-01017* / *ORA-28000* / *ORA-28001* | Contraseña incorrecta / cuenta bloqueada / contraseña vencida: la misma de SQL Developer; si venció, cambiarla ahí. |
| *DPY-3016 … cryptography package cannot be imported* | Pasaba en el ejecutable 3.3: descargar la 3.3.1 (ya trae `cryptography`). |
| No se sabe qué falla en un equipo | Doble clic en `Revisar_Equipo.bat` y mandar `autoprueba.txt` al responsable de la app. |
| La app se cierra o un botón falla | Mandar `validador.log` (junto al `.exe`) al responsable de la app. |
