# 📦 Sistema de Inventario Multi-Usuario v7.0

Sistema profesional de gestión de inventario físico con arquitectura cliente-servidor MySQL, diseñado para equipos de trabajo que realizan conteos simultáneos en bodega.

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![MySQL](https://img.shields.io/badge/MySQL-8.0+-orange.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)

---

## 🎯 Características Principales

### 📊 Gestión de Inventario
- **Conteo físico en tiempo real** de productos en bodega
- **Multi-usuario** con detección de conflictos entre equipos
- **Diferencias automáticas** entre stock sistema vs conteo físico
- **Historial completo** de movimientos y cambios
- **Novedades/comentarios** por cada ítem contado
- **Búsqueda inteligente** por código o nombre de producto
- **Filtros por líneas** de productos

### 👥 Trabajo en Equipo
- **Múltiples equipos** contando simultáneamente
- **Detección de conflictos** cuando dos equipos cuentan el mismo producto
- **Identificación de responsables** (último equipo que contó)
- **Integrantes por equipo** con registro de nombres

### 📈 Análisis y Reportes
- **KPIs en tiempo real:**
  - Porcentaje de avance del conteo
  - Productos pendientes
  - Exactitud del inventario
  - Sobrantes y faltantes
- **Export a Excel** con formato profesional:
  - Hoja completa con todos los productos
  - Hoja de diferencias con colores (verde/rojo)
  - Hoja de pendientes por contar

### 🔄 Sincronización y Rendimiento
- **Sincronización automática** cada 30 segundos (configurable)
- **Arquitectura multi-hilo** para UI fluida sin bloqueos
- **Pool de conexiones** MySQL para alto rendimiento
- **Actualización selectiva** de widgets para evitar congelamiento
- **Cache de datos** para búsquedas rápidas

### 💾 Gestión de Datos
- **Backup automático** de base de datos
- **Importación desde Excel** para crear nuevos cortes
- **Actualización de stock** desde Excel sin perder conteos
- **Reseteo completo** con backup previo
- **Restauración** de backups anteriores

### 🎨 Interfaz de Usuario
- **Diseño moderno** con CustomTkinter (tema oscuro)
- **Responsive** optimizada para resolución 1366x768
- **Tabs organizados** (Búsqueda, Pendientes, Diferencias)
- **DataGrid profesional** con columnas fijas
- **Scrollbars** en todos los paneles
- **Feedback visual y sonoro** (beeps diferenciados)
- **Consola de logs** integrada

---

## 🖥️ Capturas de Pantalla

### Panel Principal
```
┌─────────────────────────────────────────────────────────────────┐
│  SIDEBAR          │  ÁREA PRINCIPAL          │   HISTORIAL      │
│                   │                          │                  │
│  • Nuevo Corte    │  📊 KPIs (6 indicadores) │  Últimos 15      │
│  • Equipos        │  ─────────────────────   │  movimientos     │
│  • Backups        │  🔍 Búsqueda de producto │  con detalle     │
│  • Config         │  📝 Código + Cantidad    │  de equipo y     │
│  • Export Excel   │  💬 Novedad/Comentario   │  fecha/hora      │
│                   │  ─────────────────────   │                  │
│  Sesión: #123     │  📑 TABS:                │                  │
│  Equipo: A        │     • Búsqueda           │                  │
│  Filtro: 5 líneas │     • Pendientes         │                  │
│                   │     • Diferencias        │                  │
└─────────────────────────────────────────────────────────────────┘
```

---

## 🚀 Instalación

### Requisitos Previos
- **Python 3.8 o superior**
- **MySQL Server 8.0 o superior**
- **Sistema Operativo:** Windows (optimizado para Windows 10/11)

### Paso 1: Clonar el repositorio
```bash
git clone https://github.com/tu-usuario/sistema-inventario.git
cd sistema-inventario
```

### Paso 2: Instalar dependencias
```bash
pip install -r requirements.txt
```

### Paso 3: Configurar MySQL
1. Instalar y arrancar MySQL Server
2. Crear la base de datos (se crea automáticamente al iniciar)
3. Editar `config.json` si es necesario:

```json
{
  "database": {
    "host": "localhost",
    "port": 3306,
    "user": "root",
    "password": "",
    "database": "sis_inventario_db",
    "pool_size": 10
  },
  "app": {
    "sync_interval_seconds": 30
  }
}
```

### Paso 4: Ejecutar la aplicación
```bash
python inventari.py
```

---

## 📖 Guía de Uso

### 1️⃣ Crear un Nuevo Corte de Inventario
1. Click en **"Nuevo Corte"** en el sidebar
2. Asignar **nombre** a la sesión (ej: "Inventario Diciembre 2025")
3. Cargar **archivo Excel** con productos:
   - Columnas requeridas: `CODIGO`, `PRODUCTO`, `LINEA`, `STOCK`, `BODEGA`
4. Seleccionar **líneas** a inventariar (opcional)
5. El sistema crea automáticamente equipos desde el Excel

### 2️⃣ Seleccionar Equipo
1. En el sidebar, seleccionar **equipo** del dropdown
2. Cada estación debe usar un equipo diferente

### 3️⃣ Contar Productos
1. **Escanear/escribir código** y presionar ENTER
2. El sistema muestra:
   - Nombre del producto
   - Stock en sistema
   - Conteo previo (si existe)
   - Equipo que contó previamente
3. **Ingresar cantidad** contada
4. **(Opcional)** Agregar **novedad/comentario**
5. Presionar **ENTER** para guardar

### 4️⃣ Resolver Conflictos
Si otro equipo ya contó el producto:
- **SUMAR**: Agregar a la cantidad existente
- **REEMPLAZAR**: Sobrescribir el conteo anterior

### 5️⃣ Monitorear Avance
Los **KPIs** se actualizan automáticamente:
- **Avance %**: Porcentaje completado
- **Pendientes**: Productos sin contar
- **Exactitud %**: Productos con conteo exacto
- **Faltantes**: Productos con menos cantidad
- **Sobrantes**: Productos con más cantidad
- **Total**: Total de productos con stock

### 6️⃣ Exportar Resultados
1. Click en **"Export Excel"**
2. Se genera archivo con 3 hojas:
   - **Completo**: Todos los productos
   - **Diferencias**: Solo productos con diferencias (coloreado)
   - **Pendientes**: Productos sin contar

---

## 🗂️ Estructura de la Base de Datos

### Tablas Principales

#### `sesiones`
Cortes de inventario
```sql
id, nombre, fecha_inicio, fecha_fin, activo, bodega
```

#### `items_corte`
Productos de cada sesión
```sql
id, sesion_id, codigo, producto, linea, 
stock_sistema, conteo_fisico, diferencia (VIRTUAL),
novedad, fecha_conteo, ultimo_equipo_id
```

#### `equipos`
Equipos de trabajo
```sql
id, nombre_equipo, integrantes, activo, fecha_creacion
```

#### `historial_movimientos`
Auditoría de cambios
```sql
id, sesion_id, item_codigo, equipo_id, tipo_accion,
cantidad_anterior, cantidad_resultante, fecha_movimiento
```

---

## ⚙️ Configuración Avanzada

### Intervalo de Sincronización
Cambiar en **Configuración** (rango: 10-120 segundos):
```json
{
  "app": {
    "sync_interval_seconds": 30
  }
}
```

### Pool de Conexiones MySQL
Ajustar según carga de usuarios:
```json
{
  "database": {
    "pool_size": 10,
    "pool_name": "inventario_pool"
  }
}
```

### Actualización de Tabs
Los tabs se actualizan cada **3 ciclos de sincronización** para optimizar rendimiento:
- Con 30s de intervalo = actualización cada 90s
- Los datos se refrescan en cache en cada ciclo

---

## 🏗️ Arquitectura Técnica

### Stack Tecnológico
- **Frontend:** CustomTkinter (GUI moderna en Python)
- **Backend:** Python 3.8+ con threading
- **Base de Datos:** MySQL 8.0+ con connection pooling
- **Data Processing:** Pandas + OpenPyXL

### Arquitectura Multi-Hilo
```
┌─────────────────────────────────────────┐
│         Main Thread (UI)                │
│  • Renderizado de interfaz              │
│  • Eventos de usuario                   │
│  • Actualizaciones visuales             │
└─────────────────────────────────────────┘
              ↕ (after)
┌─────────────────────────────────────────┐
│    Background Thread (Sync Loop)        │
│  • Sincronización automática cada 30s   │
│  • sleep() para no bloquear             │
└─────────────────────────────────────────┘
              ↕
┌─────────────────────────────────────────┐
│  Background Threads (DB Operations)     │
│  • Consultas a MySQL                    │
│  • Guardado de conteos                  │
│  • Carga de pendientes/diferencias      │
└─────────────────────────────────────────┘
              ↕
┌─────────────────────────────────────────┐
│      MySQL Connection Pool              │
│  • 10 conexiones concurrentes           │
│  • Auto-reconexión                      │
└─────────────────────────────────────────┘
```

### Optimizaciones Clave
1. **Protección contra sync concurrente** con flag `sync_in_progress`
2. **Actualización selectiva** de widgets pesados (cada 3 ciclos)
3. **Cache de datos** en memoria para búsquedas instantáneas
4. **Operaciones DB en background** con callbacks a UI
5. **Consultas consolidadas** en un solo thread por ciclo

---

## 📁 Estructura del Proyecto

```
sistema-inventario/
│
├── inventari.py          # Aplicación principal (3000+ líneas)
├── config.json           # Configuración de DB y app
├── requirements.txt      # Dependencias Python
├── README.md            # Este archivo
│
├── logs/                # Logs diarios de operación
│   └── inventario_YYYYMMDD.log
│
├── BACKUPS_INVENTARIO/  # Backups automáticos de DB
│   └── backup_YYYYMMDD_HHMMSS.sql
│
└── exports/             # Archivos Excel exportados
    └── inventario_SESION_YYYYMMDD.xlsx
```

---

## 🐛 Solución de Problemas

### La aplicación se congela
✅ **Solucionado en v7.0**
- Threading optimizado para todas las operaciones DB
- Actualización selectiva de tabs
- Protección contra sincronizaciones concurrentes

### Error de conexión a MySQL
```
Error: No se pudo conectar a MySQL
```
**Solución:**
1. Verificar que MySQL Server esté corriendo
2. Revisar credenciales en `config.json`
3. Verificar firewall/puerto 3306

### Productos duplicados en conteo
**Solución:**
- El sistema detecta automáticamente duplicados
- Ofrece opciones: SUMAR o REEMPLAZAR
- Revisa el historial para auditoría

### Excel no se importa correctamente
**Requisitos del archivo:**
- Formato `.xlsx`
- Hoja activa con nombre específico o primera hoja
- Columnas: `CODIGO`, `PRODUCTO`, `LINEA`, `STOCK`, `BODEGA`

---

## 🔐 Seguridad

- ✅ Historial completo de auditoría
- ✅ Identificación de usuario por equipo
- ✅ Backups automáticos antes de operaciones críticas
- ✅ Validación de datos en entrada
- ✅ Transacciones MySQL para integridad de datos
- ✅ Logs detallados de todas las operaciones

---

## 🚧 Roadmap

### Próximas Funcionalidades
- [ ] Autenticación de usuarios individual
- [ ] Reportes personalizados con filtros avanzados
- [ ] Gráficos de avance en tiempo real
- [ ] Modo offline con sincronización posterior
- [ ] App móvil para conteo (Android/iOS)
- [ ] Integración con sistemas ERP
- [ ] API REST para integraciones

---

## 🤝 Contribuciones

Las contribuciones son bienvenidas. Por favor:

1. Fork el proyecto
2. Crea una rama para tu feature (`git checkout -b feature/NuevaFuncionalidad`)
3. Commit tus cambios (`git commit -m 'Agregar nueva funcionalidad'`)
4. Push a la rama (`git push origin feature/NuevaFuncionalidad`)
5. Abre un Pull Request

---

## 📄 Licencia

Este proyecto está bajo la Licencia MIT - ver el archivo [LICENSE](LICENSE) para detalles.

---

## 👨‍💻 Autor

**Sistema desarrollado para Guayas Tec**

- 📧 Email: contacto@ejemplo.com
- 🌐 Website: www.ejemplo.com

---

## 🙏 Agradecimientos

- [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) por la excelente librería de UI
- [MySQL](https://www.mysql.com/) por la robusta base de datos
- [Pandas](https://pandas.pydata.org/) por el procesamiento de datos
- La comunidad Python por las herramientas open source

---

## 📊 Estadísticas del Proyecto

- **Líneas de código:** ~3,000
- **Módulos:** 1 archivo principal modular
- **Tablas DB:** 4 tablas principales
- **Ventanas/Diálogos:** 8 ventanas diferentes
- **Hilos concurrentes:** 1-5 según carga
- **Resolución mínima:** 1366x768

---

## 🎓 Casos de Uso

### Retail
- Inventarios cíclicos mensuales
- Conteos de fin de año
- Auditorías de stock

### Manufactura
- Inventarios de materia prima
- Conteos de producto terminado
- Control de WIP (Work In Progress)

### Distribución
- Inventarios de bodegas múltiples
- Verificación de recepciones
- Conteos de despachos

---

## ⚡ Rendimiento

- **Tiempo de respuesta:** < 100ms para búsquedas
- **Usuarios simultáneos:** Hasta 50 equipos
- **Productos:** Probado con 50,000+ SKUs
- **Sincronización:** 30 segundos por defecto
- **Memoria:** ~100-150 MB en uso normal

---

**✨ Sistema robusto, escalable y fácil de usar para inventarios multi-usuario ✨**
