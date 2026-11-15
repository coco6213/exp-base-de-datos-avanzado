-- E-06c | Seguridad por roles: el analista NO puede ver el email del cliente
USE DistribuidoraAndina;
EXECUTE AS USER = 'usr_analista';
SELECT TOP 3 ClienteID, RazonSocial, Segmento FROM dbo.Cliente;   -- permitido
SELECT TOP 3 Email FROM dbo.Cliente;                              -- DENEGADO
REVERT;
