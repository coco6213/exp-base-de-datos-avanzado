/* =====================================================================
   SCRIPT 04: Seguridad - 3 roles con permisos diferentes + usuarios de prueba.
   EN SIMPLE: cada persona solo puede hacer lo que su trabajo necesita
   (principio de "minimo privilegio").
     - rol_administrador : control total de la base
     - rol_analista      : solo LEE datos para reportes (no ve email/telefono)
     - rol_operativo     : registra ventas e inventario solo usando los procedimientos
   NOTA: las contrasenas son de PRUEBA de laboratorio; cambiarlas en produccion.
   ===================================================================== */
USE master;
GO
IF SUSER_ID('login_admin_andina')     IS NULL CREATE LOGIN login_admin_andina     WITH PASSWORD = 'Lab_Admin#2026',    CHECK_POLICY = ON;
IF SUSER_ID('login_analista_andina')  IS NULL CREATE LOGIN login_analista_andina  WITH PASSWORD = 'Lab_Analista#2026', CHECK_POLICY = ON;
IF SUSER_ID('login_operativo_andina') IS NULL CREATE LOGIN login_operativo_andina WITH PASSWORD = 'Lab_Operat#2026',   CHECK_POLICY = ON;
GO
USE DistribuidoraAndina;
GO
-- Usuarios de base de datos (uno por login)
IF USER_ID('usr_admin')     IS NULL CREATE USER usr_admin     FOR LOGIN login_admin_andina;
IF USER_ID('usr_analista')  IS NULL CREATE USER usr_analista  FOR LOGIN login_analista_andina;
IF USER_ID('usr_operativo') IS NULL CREATE USER usr_operativo FOR LOGIN login_operativo_andina;
GO
-- Roles
IF DATABASE_PRINCIPAL_ID('rol_administrador') IS NULL CREATE ROLE rol_administrador;
IF DATABASE_PRINCIPAL_ID('rol_analista')      IS NULL CREATE ROLE rol_analista;
IF DATABASE_PRINCIPAL_ID('rol_operativo')     IS NULL CREATE ROLE rol_operativo;
GO
-- Administrador: control total
GRANT CONTROL ON DATABASE::DistribuidoraAndina TO rol_administrador;

-- Analista: lectura de tablas de negocio, pero sin datos personales del cliente
GRANT SELECT ON dbo.Venta        TO rol_analista;
GRANT SELECT ON dbo.DetalleVenta TO rol_analista;
GRANT SELECT ON dbo.Producto     TO rol_analista;
GRANT SELECT ON dbo.Categoria    TO rol_analista;
GRANT SELECT ON dbo.Region       TO rol_analista;
GRANT SELECT ON dbo.Sede         TO rol_analista;
GRANT SELECT ON dbo.Inventario   TO rol_analista;
GRANT SELECT ON dbo.Cliente(ClienteID, RazonSocial, Segmento, RegionID) TO rol_analista;  -- permiso por columna
DENY  SELECT ON dbo.Cliente(Email, Telefono, RUC) TO rol_analista;                        -- datos personales bloqueados
DENY  SELECT ON dbo.AuditoriaLog TO rol_analista;

-- Operativo: solo ejecuta procedimientos (no puede tocar tablas directamente)
GRANT EXECUTE ON dbo.usp_RegistrarVenta       TO rol_operativo;
GRANT EXECUTE ON dbo.usp_ActualizarInventario TO rol_operativo;
GRANT EXECUTE ON dbo.usp_GenerarAlertasStock  TO rol_operativo;
GRANT SELECT  ON dbo.Producto                 TO rol_operativo;
DENY  DELETE  ON SCHEMA::dbo                  TO rol_operativo;
GO
ALTER ROLE rol_administrador ADD MEMBER usr_admin;
ALTER ROLE rol_analista      ADD MEMBER usr_analista;
ALTER ROLE rol_operativo     ADD MEMBER usr_operativo;
GO
-- Evidencia: matriz de permisos por rol
SELECT r.name AS Rol, p.permission_name AS Permiso, p.state_desc AS Estado,
       ISNULL(OBJECT_NAME(p.major_id), p.class_desc) AS Objeto,
       ISNULL(c.name, '') AS Columna
FROM sys.database_permissions p
JOIN sys.database_principals r ON r.principal_id = p.grantee_principal_id
LEFT JOIN sys.columns c ON c.object_id = p.major_id AND c.column_id = p.minor_id
WHERE r.name LIKE 'rol_%'
ORDER BY r.name, Objeto, Columna;
GO
