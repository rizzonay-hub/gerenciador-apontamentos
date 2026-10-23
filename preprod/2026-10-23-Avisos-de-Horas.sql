/* ============================================================================
   Apontamentos de horas — Avisos ao gestor (23/10/2026)

   Decisão dos usuários: não há trava de horas. Quem passa de 8h00 (ou de 10h00)
   num dia grava normalmente, é avisado na tela, e o gestor dele recebe um aviso
   que vê no Gerenciador (botão "Avisos" na lista e aba Avisos do Gerenciamento).

   Uma linha por (pessoa, dia, gestor); se a pessoa lança mais no mesmo dia, a
   linha é atualizada e volta a "não visto".

   Quem é o gestor (ModApontamento.GestoresDe, regra em AvisoHoras.EscolherGestores):
     1. líder das equipes ativas da pessoa (TbEquipe.LIDERID, quando as equipes existirem);
     2. senão, quem tem FUNCAO com "GESTOR" no mesmo DEPTO da pessoa (TbSenha);
     3. senão, a diretoria;
     4. senão, qualquer gestor ativo.

   Rodar primeiro no Grupo_Mecano_Testes; em produção só depois do roteiro
   do IMPLANTACAO.md. Pode rodar mais de uma vez.
   ============================================================================ */

IF OBJECT_ID('dbo.TbTimeAviso', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.TbTimeAviso (
        CODIGOID   INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_TbTimeAviso PRIMARY KEY,
        USUARIOID  INT            NOT NULL,          -- quem apontou (TbSenha.CODIGOID)
        USUARIO    NVARCHAR(50)   NOT NULL,          -- nome (cópia)
        DATA       SMALLDATETIME  NOT NULL,          -- o dia apontado
        HORAS      DECIMAL(6,2)   NOT NULL,          -- total do dia depois de gravar
        LIMITE     DECIMAL(4,2)   NOT NULL,          -- 8,00 ou 10,00: o que foi ultrapassado
        GESTORID   INT            NOT NULL,          -- quem recebe o aviso (TbSenha.CODIGOID)
        GESTOR     NVARCHAR(50)   NOT NULL,          -- nome (cópia)
        CRIADOEM   DATETIME       NOT NULL CONSTRAINT DF_TbTimeAviso_CRIADOEM DEFAULT (GETDATE()),
        VISTOEM    DATETIME       NULL,              -- quando o gestor marcou como visto
        VISTOPOR   NVARCHAR(50)   NULL,
        CONSTRAINT UQ_TbTimeAviso UNIQUE (USUARIOID, DATA, GESTORID)
    );
    CREATE INDEX IX_TbTimeAviso_Gestor ON dbo.TbTimeAviso (GESTORID, VISTOEM, DATA);
    PRINT 'TbTimeAviso criada.';
END
ELSE
    PRINT 'TbTimeAviso já existia; nada feito.';
GO

/* Conferência */
SELECT GESTOR, NaoVistos = SUM(CASE WHEN VISTOEM IS NULL THEN 1 ELSE 0 END), Total = COUNT(*)
FROM dbo.TbTimeAviso GROUP BY GESTOR ORDER BY GESTOR;
