# 📦 Sistema de Gestión de Inventarios y Almacén

Este repositorio contiene una solución integral de escritorio desarrollada en Python para la administración de activos, herramientas y consumibles. El sistema automatiza el control de existencias, la generación de reportes profesionales y la gestión de préstamos, optimizando la cadena de suministro interna.

## 🛠️ ¿Qué problemas soluciona?
La gestión manual de almacenes suele derivar en pérdidas de material y datos inexactos. Este software soluciona:
* **Control de Existencias en Tiempo Real:** Elimina la incertidumbre sobre el stock mediante un registro dinámico de entradas, salidas y ubicaciones (gabinetes/cajones).
* **Falta de Trazabilidad:** Registra quién tiene cada herramienta o material a través de un módulo dedicado de préstamos y devoluciones.
* **Burocracia en Reportes:** Automatiza la creación de documentación técnica y administrativa, exportando datos a formatos estándar como Excel, Word y PDF con un solo clic.
* **Riesgo de Desabasto:** Permite la visualización rápida de cantidades críticas para asegurar la continuidad operativa.

## 🚀 Tecnologías Utilizadas
* **Lenguaje:** Python 3.x.
* **Interfaz Gráfica (GUI):** Tkinter con diseño personalizado y menús laterales.
* **Persistencia de Datos:** JSON (para almacenamiento ligero y portable).
* **Generación de Documentos:** * `openpyxl` (Reportes de inventario en Excel).
    * `python-docx` (Fichas técnicas en Word).
    * `fpdf` (Fichas de control en PDF).
* **Gestión de Archivos:** `shutil` y `os` para el manejo de rutas y copias de seguridad.

## 📊 Funcionalidades Principales
1. **Dashboard de Gestión:** Panel central para visualizar, agregar, modificar y eliminar artículos de forma intuitiva.
2. **Módulo de Préstamos:** Sistema para asignar herramientas a personal específico, manteniendo un historial de responsables.
3. **Buscador Inteligente:** Filtros por nombre, código o ubicación para agilizar la localización de materiales en almacenes grandes.
4. **Exportación Multi-formato:** - **Excel:** Listado completo de inventario para análisis de datos.
    - **Word/PDF:** Fichas técnicas individuales listas para imprimir o archivar.

## ⚙️ ¿Qué hace el sistema?
El software funciona como una estación central de control para el inventario físico, permitiendo realizar las siguientes operaciones de manera automatizada:

* **Gestión Integral de Artículos:** Permite el registro completo de productos incluyendo nombre, código único, cantidad disponible y ubicación específica (Gabinete/Cajón) dentro del almacén.
* **Control de Stock Dinámico:** Facilita la actualización inmediata de existencias (entradas y salidas) y permite la edición o eliminación de registros para mantener la base de datos depurada.
* **Administración de Préstamos:** Gestiona la asignación temporal de herramientas o materiales a empleados, vinculando cada artículo con un responsable para asegurar su devolución.
* **Búsqueda y Filtrado Inteligente:** Implementa un motor de búsqueda que localiza artículos en tiempo real por diversos criterios, agilizando la consulta en inventarios extensos.
* **Automatización de Documentación (Reportes):** * **Genera Reportes en Excel:** Crea una hoja de cálculo profesional con el inventario completo para auditorías o análisis financiero.
    * **Crea Fichas Técnicas en Word/PDF:** Produce documentos individuales con el logo de la institución y los detalles del producto, listos para impresión o archivo digital.
* **Persistencia de Datos Segura:** Utiliza un sistema de archivos JSON que guarda automáticamente la información al cerrar el programa, garantizando que no haya pérdida de datos entre sesiones.
* **Interfaz de Usuario Intuitiva:** Despliega una ventana organizada con tablas visuales, botones de acción rápida y cuadros de diálogo de confirmación para minimizar errores operativos.

## 📖 Manual de Uso
1. **Ejecución:** Inicie el programa ejecutando `Sistema de Inventario (Almacen).py`.
2. **Registro:** Utilice el botón "Nuevo" para dar de alta productos, asignando códigos únicos y ubicaciones físicas.
3. **Mantenimiento:** Seleccione cualquier registro de la tabla para modificar stock o generar sus fichas técnicas.
4. **Reportes:** Acceda a los botones de exportación en la barra lateral para generar los informes necesarios.

## 👥 Desarrollador
* **Ramirez Cardenas Luis Armando** - (Matrícula: 2200607)

**Institución:** Universidad Autónoma de Baja California (UABC).
**Facultad:** Contaduría y Administración.
**Carrera:** Inteligencia de Negocios.
