-- ============================================
-- BANCO: Marketplace (SQL Server)
-- Versão T-SQL do schema.sql (SQLite) da raiz do projeto.
-- Mesma modelagem, adaptada aos tipos/sintaxe do SQL Server.
-- Rode este script já conectado ao banco "Marketplace"
-- (o importar.py cria o banco antes de rodar este script).
-- ============================================

-- ============================================
-- CANAIS
-- Cadastro de referência: canal (marketplace, varejista parceiro ou
-- instituição financeira) + loja (unidade/CNPJ) + vertente (1P/3P/SEM ID).
-- Extraído e normalizado a partir da aba Visão_AGOSTO_2026 do arquivo
-- "Calendário_Planejamento Baixas_2026.xlsx" (76 combinações).
-- Vertente é estável por canal+loja (sem conflitos observados nos dados);
-- banco/conta/agência NÃO são estáveis (mudaram ao longo do tempo para o
-- mesmo canal+loja), por isso ficam em conciliacao_diaria, não aqui.
-- ============================================

IF OBJECT_ID('dbo.canais', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.canais (
        id       INT IDENTITY(1,1) PRIMARY KEY,
        canal    NVARCHAR(100) NOT NULL,
        loja     NVARCHAR(100) NOT NULL,
        vertente NVARCHAR(10) NOT NULL DEFAULT 'SEM ID' CHECK (vertente IN ('1P', '3P', 'SEM ID')),
        CONSTRAINT uq_canais_canal_loja UNIQUE (canal, loja)
    );
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_canais_canal' AND object_id = OBJECT_ID('dbo.canais'))
    CREATE INDEX idx_canais_canal ON dbo.canais(canal);
GO

-- ============================================
-- CONCILIACAO_DIARIA
-- Uma linha por canal + data (equivalente à aba "Visão_MÊS").
-- valor_eb = Extrato Bancário (manual)
-- valor_ab = Arquivo de Baixa, o repasse relatado pelo próprio marketplace
--            (auto quando existir tabela mkp_<canal> com import; manual
--            quando não houver import automático pra esse canal)
-- valor_rb = valor registrado no sistema interno (manual)
-- ============================================

IF OBJECT_ID('dbo.conciliacao_diaria', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.conciliacao_diaria (
        id            INT IDENTITY(1,1) PRIMARY KEY,
        canal_id      INT NOT NULL REFERENCES dbo.canais(id),
        data          DATE NOT NULL,
        banco         NVARCHAR(100),
        conta         NVARCHAR(50),
        agencia       NVARCHAR(50),
        valor_eb      DECIMAL(18,2),
        valor_ab      DECIMAL(18,2),
        valor_rb      DECIMAL(18,2),
        diferenca     AS (ROUND(ISNULL(valor_eb, 0) - ISNULL(valor_ab, 0), 2)) PERSISTED,
        observacao    NVARCHAR(MAX),
        responsavel   NVARCHAR(100),
        feito_ra      BIT NOT NULL DEFAULT 0,
        status_ra     NVARCHAR(20) NOT NULL DEFAULT 'PENDENTE' CHECK (status_ra IN ('PENDENTE', 'CONCILIADO')),
        previsto      DECIMAL(18,2),
        comissao      DECIMAL(18,2),
        perc_comissao DECIMAL(9,4),
        notas_fiscais NVARCHAR(MAX),
        atualizado_em DATETIME2 NOT NULL DEFAULT SYSDATETIME(),
        CONSTRAINT uq_conciliacao_canal_data UNIQUE (canal_id, data)
    );
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_conciliacao_data' AND object_id = OBJECT_ID('dbo.conciliacao_diaria'))
    CREATE INDEX idx_conciliacao_data ON dbo.conciliacao_diaria(data);
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_conciliacao_status' AND object_id = OBJECT_ID('dbo.conciliacao_diaria'))
    CREATE INDEX idx_conciliacao_status ON dbo.conciliacao_diaria(status_ra);
GO

-- ============================================
-- MKP_MAGALU
-- Dado transacional detalhado (pedido/parcela) do Magalu, equivalente à
-- aba EMPILHADO. Colunas normalizadas a partir de:
--  - histórico já consolidado (aba EMPILHADO, 27 colunas)
--  - export bruto periódico do marketplace (Sheet1, 30 colunas)
-- Sem UNIQUE de negócio nas colunas (id_transacao/external_withdraw_id
-- NÃO são únicos por linha: o mesmo id_transacao repete entre pedido e
-- estorno, e external_withdraw_id é um lote de saque que agrupa milhares
-- de linhas). A deduplicação de reimport é feita a nível de ARQUIVO, veja
-- importacoes_log.
-- ============================================

IF OBJECT_ID('dbo.mkp_magalu', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.mkp_magalu (
        id                           INT IDENTITY(1,1) PRIMARY KEY,
        loja                         NVARCHAR(100) NOT NULL,
        data_repasse                 DATE,
        data_transacao               DATE,
        data_pedido                  DATE,
        id_pedido                    NVARCHAR(100),
        pedido_mkp                   NVARCHAR(100),
        id_transacao                 NVARCHAR(100),
        numero_nota_fiscal           NVARCHAR(50),
        nome_cliente                 NVARCHAR(200),
        metodo_pagamento             NVARCHAR(50),
        parcela_atual                INT,
        total_parcelas               INT,
        valor_repasse                DECIMAL(18,2),
        valor_antecipacao            DECIMAL(18,2),
        perc_taxa_antecipacao        DECIMAL(9,4),
        comissao_por_parcela         DECIMAL(18,2),
        valor_bruto_parcela          DECIMAL(18,2),
        valor_bruto_pedido           DECIMAL(18,2),
        perc_comissao                DECIMAL(9,4),
        valor_comissao               DECIMAL(18,2),
        valor_bruto_seller           DECIMAL(18,2),
        origem                       NVARCHAR(100),
        observacoes                  NVARCHAR(MAX),
        external_withdraw_id         NVARCHAR(100),
        valor_fixo_pedido_parcelada  DECIMAL(18,2),
        valor_fixo_pedido            DECIMAL(18,2),
        arquivo_origem               NVARCHAR(255) NOT NULL,
        importado_em                 DATETIME2 NOT NULL DEFAULT SYSDATETIME()
    );
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_mkp_magalu_data_repasse' AND object_id = OBJECT_ID('dbo.mkp_magalu'))
    CREATE INDEX idx_mkp_magalu_data_repasse ON dbo.mkp_magalu(data_repasse);
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_mkp_magalu_loja' AND object_id = OBJECT_ID('dbo.mkp_magalu'))
    CREATE INDEX idx_mkp_magalu_loja ON dbo.mkp_magalu(loja);
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_mkp_magalu_external_withdraw' AND object_id = OBJECT_ID('dbo.mkp_magalu'))
    CREATE INDEX idx_mkp_magalu_external_withdraw ON dbo.mkp_magalu(external_withdraw_id);
GO

-- ============================================
-- CALENDARIO_BAIXAS
-- Uma linha por canal + data, empilhado a partir das abas "Visão_MÊS" do
-- Calendario.xlsx (só o valor_ab; EB/RB e o resto ficam de fora). ano/mes
-- separados de "data" pra filtrar/agrupar por período sem função de data.
-- canal_id é o Id pra cruzar com canais (canal/loja/vertente).
-- ============================================

IF OBJECT_ID('dbo.calendario_baixas', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.calendario_baixas (
        id            INT IDENTITY(1,1) PRIMARY KEY,
        canal_id      INT NOT NULL REFERENCES dbo.canais(id),
        ano           INT NOT NULL,
        mes           INT NOT NULL,
        data          DATE NOT NULL,
        valor_ab      DECIMAL(18,2) NOT NULL,
        atualizado_em DATETIME2 NOT NULL DEFAULT SYSDATETIME(),
        CONSTRAINT uq_calendario_baixas_canal_data UNIQUE (canal_id, data)
    );
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_calendario_baixas_ano_mes' AND object_id = OBJECT_ID('dbo.calendario_baixas'))
    CREATE INDEX idx_calendario_baixas_ano_mes ON dbo.calendario_baixas(ano, mes);
GO

-- ============================================
-- IMPORTACOES_LOG
-- Controla quais arquivos já foram importados (por hash de conteúdo),
-- pra evitar empilhar a mesma base duas vezes.
-- ============================================

IF OBJECT_ID('dbo.importacoes_log', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.importacoes_log (
        id           INT IDENTITY(1,1) PRIMARY KEY,
        canal        NVARCHAR(100) NOT NULL,
        arquivo      NVARCHAR(255) NOT NULL,
        hash_arquivo CHAR(64) NOT NULL UNIQUE,
        linhas       INT NOT NULL,
        importado_em DATETIME2 NOT NULL DEFAULT SYSDATETIME()
    );
END;
GO

-- ============================================
-- VIEW: vw_conciliacao_status
-- Reproduz a "Visão_MÊS", já com o status calculado a partir da batida
-- EB x AB x RB (além do status_ra, que continua sendo controle manual
-- da equipe).
-- ============================================

CREATE OR ALTER VIEW dbo.vw_conciliacao_status AS
SELECT
    cd.id,
    c.canal,
    c.loja,
    c.vertente,
    cd.data,
    cd.banco,
    cd.conta,
    cd.agencia,
    cd.valor_eb,
    cd.valor_ab,
    cd.valor_rb,
    cd.diferenca,
    CASE
        WHEN cd.valor_eb IS NULL OR cd.valor_ab IS NULL OR cd.valor_rb IS NULL THEN 'INCOMPLETO'
        WHEN cd.valor_eb = cd.valor_ab AND cd.valor_ab = cd.valor_rb THEN 'CONCILIADO'
        ELSE 'DIVERGENTE'
    END AS status_calculado,
    cd.status_ra,
    cd.feito_ra,
    cd.responsavel,
    cd.observacao,
    cd.previsto,
    cd.comissao,
    cd.perc_comissao,
    cd.notas_fiscais
FROM dbo.conciliacao_diaria cd
JOIN dbo.canais c ON c.id = cd.canal_id;
GO

-- ============================================
-- SEED: canais (76 combinações canal+loja+vertente, normalizadas a
-- partir da aba Visão_AGOSTO_2026). Idempotente via MERGE.
-- 'SICRED' e 'SICREDI' apareciam como canais distintos na planilha
-- original por grafia divergente; confirmado que é a mesma cooperativa,
-- unificado em 'SICREDI'.
-- ============================================

MERGE dbo.canais AS destino
USING (VALUES
    ('AMAZON', 'CENTER', '1P'),
    ('AMAZON', 'FULL', '1P'),
    ('AMAZON', 'LG', '3P'),
    ('AMAZON', 'WEB', '3P'),
    ('ANGELONI', 'LG', '3P'),
    ('ANGELONI', 'WEB', '3P'),
    ('BANCO INTER', 'CENTER', '1P'),
    ('BANCO INTER', 'LG', '3P'),
    ('BANCO INTER', 'WEB', '3P'),
    ('BARRA CRED', 'CENTER', '1P'),
    ('BARRA CRED', 'WEB', '3P'),
    ('BRADESCOSHOP', 'WEB', '3P'),
    ('CAMICADO', 'WEB/LG', 'SEM ID'),
    ('CARREFOUR', 'CENTER', '1P'),
    ('CARREFOUR', 'LG', '3P'),
    ('CARREFOUR', 'SEM ID', 'SEM ID'),
    ('CARREFOUR', 'WEB', '3P'),
    ('CASA E VIDEO', 'LG', '3P'),
    ('CASA E VIDEO', 'SEM ID', 'SEM ID'),
    ('COLOMBO', 'CENTER', '1P'),
    ('DECATHLON', 'CENTER', '1P'),
    ('DROGA RAIA', 'SEM ID', 'SEM ID'),
    ('E-FÁCIL', 'CENTER', '1P'),
    ('E-FÁCIL', 'LG', '3P'),
    ('E-FÁCIL', 'SEM ID', 'SEM ID'),
    ('FASTSHOP', 'SEM ID', 'SEM ID'),
    ('IMPERIO MOVEIS', 'CENTER', '1P'),
    ('KABUM', 'CENTER', '1P'),
    ('KABUM', 'WEB', '3P'),
    ('KOERICK', 'SEM ID', 'SEM ID'),
    ('LE BISCUIT', 'SEM ID', 'SEM ID'),
    ('LEROY', 'CENTER', '1P'),
    ('LEROY', 'LG', '3P'),
    ('LEROY', 'WEB', '3P'),
    ('LEROY', 'WEB/LG', '3P'),
    ('LIVELO', 'CENTER', '1P'),
    ('LIVELO', 'LG', '3P'),
    ('LIVELO', 'WEB', '3P'),
    ('LL LOYALT', 'CENTER', '1P'),
    ('LL LOYALT', 'WEB', '3P'),
    ('MADEIRA MADEIRA', 'WEB', '3P'),
    ('MAGALU', 'CENTER', '1P'),
    ('MAGALU', 'COLETAS ES', '1P'),
    ('MAGALU', 'WEB', '3P'),
    ('MAIS CORREIOS', 'CENTER', '1P'),
    ('MERCADO', 'CENTER', '1P'),
    ('MERCADO', 'COLETAS 05', '1P'),
    ('MERCADO', 'COLETAS ES', '1P'),
    ('MERCADO', 'ELETRO', '3P'),
    ('MERCADO', 'FULL', '1P'),
    ('MERCADO', 'MOVEIS', '3P'),
    ('MERCADO', 'PRIME', '3P'),
    ('MERCADO', 'WEB', '3P'),
    ('NETSHOES', 'CENTER', '1P'),
    ('PAGAR-ME', 'SEM ID', 'SEM ID'),
    ('PICPAY', 'CENTER', '1P'),
    ('QUERO QUERO', 'WEB', '3P'),
    ('RI HAPPY', 'CENTER', '1P'),
    ('SENFF', 'SEM ID', 'SEM ID'),
    ('SHOP HUB', 'CENTER', '1P'),
    ('SHOP HUB', 'WEB', '3P'),
    ('SHOPEE', 'CENTER', '1P'),
    ('SHOPEE', 'ES', '1P'),
    ('SHOPEE', 'PRIME', '3P'),
    ('SHOPEE', 'WEB', '3P'),
    ('SHOPPING BB', 'SEM ID', 'SEM ID'),
    ('SICOOB', 'CENTER', '1P'),
    ('SICOOB', 'WEB', '3P'),
    ('SICOOB', 'WEB/LG', '3P'),
    ('SICREDI', 'CENTER', '1P'),
    ('TIKTOK', 'CENTER', '1P'),
    ('TIKTOK', 'COLETAS', '1P'),
    ('VIA VAREJO', 'CENTER', '1P'),
    ('VIA VAREJO', 'WEB', '3P'),
    ('ZEMA', 'WEB', '3P'),
    ('ZOOM', 'CENTER', '1P')
) AS origem (canal, loja, vertente)
ON destino.canal = origem.canal AND destino.loja = origem.loja
WHEN NOT MATCHED THEN
    INSERT (canal, loja, vertente) VALUES (origem.canal, origem.loja, origem.vertente);
GO
