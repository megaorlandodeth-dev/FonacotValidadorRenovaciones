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
   - **🌐 Modo Concentrado General del Día:** Para quien genera el lote final diario. Consolida todos los tickets del día sin duplicar la asignación inicial de Nancy (2609) con la del agente asignado final. Detecta además duplicados de renovaciones en diferentes tickets.
   - **👤 Modo por Compañero (Gafete o Nombre):** Permite filtrar por analista específico respetando la regla de Nancy (2609) y mostrando alertas si el ticket fue reasignado entre compañeros.

3. **Herramientas de Cierre Diario:**
   - **Checkboxes de Validación (`☑️ / ☐`):** Marca renovaciones validadas con doble clic o barra espaciadora.
   - **Copiar Lote para Correo:** Copia la tabla concentrada al portapapeles (`Ctrl+V`) lista para enviar.
   - **Exportar a Excel:** Genera reporte multi-pestaña con enlaces directos a iTop y semáforo de colores.
   - **Acceso Directo a iTop:** Botón para abrir el ticket seleccionado en el navegador.

---

## 📦 Estructura del Proyecto

- `validador_renovaciones_app.py`: Interfaz gráfica moderna desarrollada en `CustomTkinter`.
- `core_renovaciones.py`: Motor de parsing, clasificación regex, normalización y desglose 1 a N de renovaciones.
- `outlook_renovaciones_engine.py`: Conector MAPI con Microsoft Outlook, escaneo multi-buzón y gestión de historial de asignaciones.
- `Ejecutar_Validador_Renovaciones.bat`: Script de inicio rápido.

---

## 🛠️ Requisitos e Instalación

1. **Python 3.10+** (probado en Python 3.12).
2. **Dependencias:**
   ```bash
   pip install customtkinter openpyxl pandas pywin32 tkcalendar babel
   ```
3. **Microsoft Outlook** con sesión activa en Windows.
4. **Ejecución:**
   ```bash
   python validador_renovaciones_app.py
   ```

