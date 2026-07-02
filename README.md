# AllRide – Cuadratura de viajes

Aplicación Streamlit para cuadrar el consolidado del cliente con los servicios de AllRide.

## Lógica

Por cada viaje del cliente (fecha + ruta + empresa):

1. **Hora AllRide** = `HORA DE POSTURA + 15 min` (si existe esa columna) o la hora directa
2. Si ya existe en AllRide con esa hora exacta → **✅ OK, conservar**
3. Si NO existe con esa hora pero hay otro horario de la misma ruta ese día → **✏️ editar el más cercano**
4. Si hay viajes en AllRide de esa ruta ese día sin match con el cliente → **❌ cancelar**
5. Si no existe ningún viaje de esa ruta ese día en AllRide → **➕ crear**

## Archivos que necesita

| Archivo | Obligatorio | Descripción |
|---------|:-----------:|-------------|
| Consolidado del cliente | ✅ | Excel con hoja RESUMEN o Programación RM |
| Exportación AllRide | ✅ | Servicios exportados desde AllRide |
| Plantilla cancelación masiva | ✅ | Archivo `Cancelación_de_servicios_...xlsx` de AllRide |
| Plantilla edición horarios | ⬜ | Archivo `Edición_Horarios_...xlsx` de AllRide (opcional) |

## Outputs generados

- **Resumen completo** — Excel con 4 hojas: OK, Editar, Cancelar, Crear
- **Cancelación masiva** — misma plantilla con X en columna Cancelar
- **Edición horarios** — plantilla con los IDs y nuevas horas
- **Viajes a crear** — listado de los que no existen en AllRide

## Instalación local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy en Streamlit Cloud (gratis)

1. Sube este repositorio a GitHub (público o privado)
2. Ve a [share.streamlit.io](https://share.streamlit.io)
3. Conecta tu cuenta GitHub
4. Selecciona el repositorio y el archivo `app.py`
5. Click **Deploy** — listo en ~2 minutos

## Columnas detectadas automáticamente

**Consolidado cliente:**
- Ruta: `NUEVO NOMBRE RUTA FINAL` → `NOMBRE RUTA FINAL` → `NOMBRE RUTA`
- Hora postura: `HORA DE POSTURA`
- Hora llegada: `HORA DE LLEGADA A BODEGA` / `HORA DE LLEGADA\n HORA DE SALIDA`
- Fecha: `FECHA`
- Empresa: `EMPRESA`
- Tipo: `TIPO DE PEDIDO` / `TIPO`

**AllRide:**
- Fecha/hora: `Fecha estimada de inicio` (fallback: `Fecha de inicio`)
- Ruta: `Ruta` (se elimina prefijo `RDD -` automáticamente)
- Empresa: `Comunidades`
- Tipo: `Tipo`
- ID: `ID de servicio` / `Servicio`
