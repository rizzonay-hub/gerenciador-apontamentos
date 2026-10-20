/* ============================================================================
   Apontamentos de horas — "Meus grupos de empresas" (20/10/2026)

   Cada pessoa monta grupos de empresas que marca sempre juntas no lançamento
   com várias empresas. Uma linha por (pessoa, grupo, empresa).

   Rodar primeiro no Grupo_Mecano_Testes; em produção só depois do roteiro
   do IMPLANTACAO.md. Pode rodar mais de uma vez (não recria o que já existe).
   Nada mais muda no banco: a TbTime e as procedures Time_* ficam como estão.
   ============================================================================ */

IF OBJECT_ID('dbo.TbTimeGrupoEmpresa', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.TbTimeGrupoEmpresa (
        CODIGOID   INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_TbTimeGrupoEmpresa PRIMARY KEY,
        USUARIOID  INT           NOT NULL,                 -- TbSenha.CODIGOID (dono do grupo)
        GRUPO      NVARCHAR(60)  NOT NULL,                 -- nome que a pessoa deu
        EMPRESAID  NVARCHAR(7)   NOT NULL,                 -- TbClientes.TRADUTOR
        CONSTRAINT UQ_TbTimeGrupoEmpresa UNIQUE (USUARIOID, GRUPO, EMPRESAID)
    );
    CREATE INDEX IX_TbTimeGrupoEmpresa_Usuario ON dbo.TbTimeGrupoEmpresa (USUARIOID, GRUPO);
    PRINT 'TbTimeGrupoEmpresa criada.';
END
ELSE
    PRINT 'TbTimeGrupoEmpresa já existia; nada feito.';
GO

/* Conferência */
SELECT USUARIOID, GRUPO, Empresas = COUNT(*) FROM dbo.TbTimeGrupoEmpresa GROUP BY USUARIOID, GRUPO ORDER BY USUARIOID, GRUPO;
