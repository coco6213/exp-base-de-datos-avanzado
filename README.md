# DataShield 360 – Distribuidora Andina Comercial S.A.C.

Proyecto del curso **Base de Datos Avanzadas y Big Data (CIIN1487R)** – avance de la semana 6:
solución de **Inteligencia de Negocios** (hasta el tema *Inteligencia de Negocios y Visualización de Datos*).

**Grupo:** Channel, Yin, Gutiérrez, Pirsin · **Versión:** `v0.6-avance-s6`

## ¿Qué hay aquí? (en simple)

| Carpeta | Contenido |
|---|---|
| `sql/` | Scripts numerados 01–11 (crear base, datos, automatización, seguridad, validación, ETL + Data Warehouse, experimento, pruebas, respaldos, vista del dashboard, reconciliación). `sql/demo/` = consultas usadas para las capturas. |
| `scripts/` | PowerShell: crea el dashboard en Excel, los gráficos del experimento, captura pantallas y llena el informe. |
| `dashboard/` | `Dashboard_Andina.xlsx` – dashboard BI conectado al Data Warehouse. |
| `diagramas/` | Arquitectura, diagrama entidad-relación y esquema estrella (HTML + PNG). |
| `evidencias/` | Capturas con código E-xx, marcadas con cuadros rojos. |
| `DATOS_FICTICIOS/` | **Los datos simulados** exportados a CSV (base diaria, Data Warehouse, prueba de viabilidad) y lo que salió de cada consulta. Empezar por `LEEME_PRIMERO.txt`. |
| `resultados/` | Logs de las ejecuciones (RUN-xx) y salidas de los scripts que no están en DATOS_FICTICIOS. |
| `informe/` | Informe de Proyecto Final (avance) y presentación. |

## Requisitos

- SQL Server 2025 (Developer) con autenticación de Windows, `sqlcmd` 18 y ODBC Driver 18.
- Microsoft Excel 365 (para el dashboard) y PowerShell 5.1.

## Cómo reproducir todo (≈ 3 minutos)

Ejecutar desde la carpeta raíz, en este orden (`-I` activa QUOTED_IDENTIFIER, `-C` confía en el certificado local):

```bash
sqlcmd -S localhost -E -C -I -i sql/01_crear_bd_oltp.sql
sqlcmd -S localhost -E -C -I -i sql/02_generar_datos.sql
sqlcmd -S localhost -E -C -I -i sql/05_validacion_dataset.sql
sqlcmd -S localhost -E -C -I -i sql/03_procedimientos_triggers.sql
sqlcmd -S localhost -E -C -I -i sql/04_seguridad_roles.sql
sqlcmd -S localhost -E -C -I -i sql/08_pruebas_calidad_seguridad.sql
sqlcmd -S localhost -E -C -I -i sql/06_dw_etl.sql
sqlcmd -S localhost -E -C -I -i sql/11_reconciliacion_etl.sql
sqlcmd -S localhost -E -C -I -i sql/09_politica_respaldo.sql
sqlcmd -S localhost -E -C -I -i sql/07_experimento.sql
sqlcmd -S localhost -E -C -I -i sql/10_vista_dashboard.sql
powershell -ExecutionPolicy Bypass -File scripts/crear_dashboard_excel.ps1
```

**Comprobación de reproducibilidad:** el script 05 muestra la huella `V0`; debe salir
`HuellaVenta = 2115345625` y `HuellaDetalle = -1892332103`.

## Resultados clave del avance

- Dataset simulado: 200 000 ventas, 599 806 líneas (2023–2025), 13 validaciones de calidad: todas PASA.
- ETL: 599 808 extraídas → 17 667 rechazadas (anuladas) → 582 141 cargadas; reconciliación S/ 0,00 de diferencia.
- Experimento (10 corridas): el Data Warehouse con índice columnar es **7,7 a 11,5 veces más rápido** que la base transaccional.
- Pruebas T-01 a T-10: todas PASA (transacciones, auditoría, roles, respaldos).

## Pendiente para la entrega final

Componente NoSQL (MongoDB), solución Big Data (PySpark), integración al dashboard e Informe de Medición.

> Las contraseñas del script 04 son solo de laboratorio. Cámbielas antes de usar la base fuera del aula.
