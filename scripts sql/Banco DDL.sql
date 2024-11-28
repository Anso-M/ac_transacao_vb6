-- CRIAR O BANCO E AS TABELAS CLIENTES E TRANSAÇÕES

CREATE DATABASE banco;

USE banco;

CREATE TABLE clientes (
    id_cliente INT IDENTITY(1, 1) PRIMARY KEY,
    nome_cliente VARCHAR(100) NOT NULL,
    numero_cartao VARCHAR(16) NOT NULL UNIQUE,
	CONSTRAINT CHK_Clientes_NumeroCartao CHECK (LEN(numero_cartao) = 16)
);

CREATE TABLE transacoes (
    id_transacao INT IDENTITY(1, 1) PRIMARY KEY,
    numero_cartao VARCHAR(16) NOT NULL,
    valor_transacao DECIMAL(10, 2) NOT NULL,
    data_transacao DATE NOT NULL,
    descricao VARCHAR(255),
    id_cliente INT,
    CONSTRAINT FK_Transacoes_Clientes FOREIGN KEY (id_cliente) REFERENCES Clientes(id_cliente),
	CONSTRAINT CHK_Transacoes_NumeroCartao CHECK (LEN(numero_cartao) = 16)
);


/* INSERIR CLIENTES NO BANCO DE DADOS PARA TESTES POSTERIORES DA INTERFACE VB6
PARA INSERÇÃO DE TRANSAÇÕES */

INSERT INTO clientes (nome_cliente, numero_cartao)
VALUES ('Ana', '0000000000000000');

INSERT INTO clientes (nome_cliente, numero_cartao)
VALUES ('Bob', '1111111111111111');