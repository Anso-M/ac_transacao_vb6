CREATE PROCEDURE sp_CalcularTotaisPorPeriodo
    @Data_Inicial DATE,
    @Data_Final DATE
AS
BEGIN
    SELECT 
        numero_cartao,
        SUM(valor_transacao) AS valor_total,
        COUNT(*) AS quantidade_transacoes
    FROM 
        transacoes
    WHERE 
        data_transacao BETWEEN @Data_Inicial AND @Data_Final
    GROUP BY 
        numero_cartao
END


-- EXECUÇÃO DA STORED PROCEDURE
EXEC sp_CalcularTotaisPorPeriodo '2008-01-01', '2021-01-01';



CREATE FUNCTION fn_CategoriaTransacao (@Valor_Transacao DECIMAL(10, 2))
RETURNS VARCHAR(10)
AS
BEGIN
    DECLARE @Categoria VARCHAR(10)

    IF @Valor_Transacao > 1000
        SET @Categoria = 'Alta'
    ELSE IF @Valor_Transacao BETWEEN 500 AND 1000
        SET @Categoria = 'Média'
    ELSE
        SET @Categoria = 'Baixa'

    RETURN @Categoria
END


-- CONSULTA USANDO A FUNÇÃO ANTERIOR
SELECT 
    id_transacao,
    numero_cartao,
    valor_transacao,
    data_transacao,
    descricao,
    dbo.fn_categoriaTransacao(valor_transacao) AS categoria
FROM 
    transacoes;



CREATE VIEW vw_TransacoesDetalhadas AS
SELECT 
    c.nome_cliente,
    t.numero_cartao,
    t.valor_transacao,
    t.data_transacao,
    dbo.fn_CategoriaTransacao(t.valor_transacao) AS categoria
FROM 
    transacoes t
INNER JOIN 
    clientes c ON t.id_cliente = c.id_cliente;


-- USO DA VIEW
SELECT * FROM vw_TransacoesDetalhadas;