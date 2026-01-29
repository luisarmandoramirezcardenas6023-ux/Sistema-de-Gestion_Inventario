#  Sistema de Gestión de Inventario Pro (Almacén-Taller)

Este es un sistema de escritorio robusto desarrollado en Python para el control de inventarios físicos en entornos de taller o almacén. El software no solo registra existencias, sino que gestiona el flujo de herramientas hacia los empleados, manteniendo un historial detallado de quién tiene qué material.

##  Características Principales
- **Gestión CRUD:** Altas, bajas y modificaciones de herramientas con validación de códigos únicos.
- **Módulo de Préstamos Masivos:** Permite seleccionar múltiples herramientas para un solo empleado en una sola transacción.
- **Semáforo de Stock:** Indicadores visuales automáticos:
  - 🔴 **Rojo:** Stock crítico (2 o menos).
  - 🟡 **Amarillo:** Stock bajo (5 o menos).
- **Exportación Multi-formato:** Generación de fichas técnicas y auditorías en **Excel, Word y PDF**.
- **Seguridad de Datos:** - Copias de seguridad automáticas al inicio del sistema.
  - Visor histórico de respaldos integrado.
- **Historial Global:** Registro cronológico de todas las acciones del sistema.

##  Soluciones que ofrece este sistema
Este software está diseñado para resolver problemáticas comunes en la gestión de activos:
1. **Pérdida de Herramientas:** Soluciona la falta de control sobre quién retiró un equipo mediante el registro por Número de Empleado.
2. **Quiebres de Stock:** Evita quedarse sin material gracias al sistema de alertas por colores.
3. **Auditorías Lentas:** Reduce horas de trabajo administrativo generando reportes de movimientos diarios con un solo clic.
4. **Falta de Trazabilidad:** Permite ver el historial completo de una sola pieza, desde su creación hasta sus múltiples entradas/salidas.
5. **Errores de Captura:** Bloquea entradas de texto en campos numéricos y evita la duplicidad de códigos de barras.

##  Requisitos
Para ejecutar este sistema, necesitas instalar las siguientes dependencias:

```bash
pip install openpyxl python-docx fpdf
