Alteração na Tabela Cobrança Fidelidade

```
ALTER TABLE SIAF_PLUS.dbo.COBRANCAS_FIDELIDADES
ADD ADESAO_ANTECIPADA bit NOT NULL Default (0);
```


ALTER TABLE SIAF_PLUS.producao.COBRANCAS_SERVICOS
ADD ADESAO_ANTECIPADA bit NOT NULL Default (0);


  return json_encode(["sucesso" => true, "redireciona" => false, "mensagem" => "Dados incorretos"]);