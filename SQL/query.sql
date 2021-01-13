

DELETE FROM SIAF_PLUS.producao.COBRANCAS_SERVICOS WHERE COD_COBRANCA= 161

select * from  SIAF_PLUS.producao.COBRANCAS_SERVICOS WHERE COD_COBRANCA = 161
select * from SIAF_PLUS.dbo.MENSALIDADES_SERVICOS MENSALIDADES_SERVICOS
 
DELETE FROM SIAF_PLUS.producao.COBRANCAS_SERVICOS WHERE COD_COBRANCA = 161

SELECT * FROM  SIAF_PLUS.producao.COBRANCAS_SERVICOS WHERE COD_COBRANCA = 161

UPDATE SIAF_PLUS.producao.COBRANCAS_SERVICOS SET DAT_FATURA_COBR_SERV = (SELECT GETDATE()-DAY(GETDATE())+1) WHERE COD_COBRANCA = 161

INSERT INTO SIAF_PLUS.producao.COBRANCAS_SERVICOS (COD_COBRANCA,COD_SERVICO,QTD_COBR_SERV,INS_COBR_SERV,VAL_COBR_SERV,VAL_PROMO_COBR_SERV, DAT_FATURA_COBR_SERV) 



 SELECT c.COD_CLIENTE, NOM_CLIENTE, END_COBRANCA, CID_COBRANCA, BAI_COBRANCA, CEP_COBRANCA, EST_COBRANCA, CPF_CLIENTE, PES_CLIENTE, COD_COBRANCA, VAL_COBRANCA, 
			  DTI_COBRANCA, DTV_COBRANCA, VAL_PROMO_COBRANCA, TAX_BNC_COBRANCA, TAX_SUP_COBRANCA 
			  FROM siaf_plus.producao.CLIENTES c INNER JOIN siaf_plus.producao.COBRANCAS o ON o.COD_CLIENTE=c.COD_CLIENTE 
			  WHERE c.COD_PROVEDOR=1 
			  and CPF_CLIENTE <> (SELECT CPF_CLIENTE FROM [SIAF_PLUS].producao.CLIENTES WHERE COD_CLIENTE = 161)
			  OR C.COD_CLIENTE = 161
			  and STA_COBRANCA IN ('ATIVO','CANCELADO','CJ - REGULARIZANDO') AND (GER_COBRANCA=1) or (STA_COBRANCA='BLOQUEADO' AND DATEDIFF(d,DTB_COBRANCA,GETDATE())<30) 			  
			  AND (VAL_COBRANCA>0 AND (VAL_PROMO_COBRANCA>0 OR (VAL_PROMO_COBRANCA=0 AND (EXP_PROMO_COBRANCA IS NULL OR EXP_PROMO_COBRANCA<=GETDATE())))) 
			  AND NOT EXISTS (SELECT COD_COBRANCA FROM siaf_plus.producao.MENSALIDADES 
			  WHERE o.COD_COBRANCA = COD_COBRANCA AND (DATEDIFF(d, GETDATE(), DAT_VENC_MENSALIDADE) >= 7))



INSERT INTO SIAF_PLUS.producao.COBRANCAS_SERVICOS (COD_COBRANCA,COD_SERVICO,QTD_COBR_SERV,INS_COBR_SERV,VAL_COBR_SERV,VAL_PROMO_COBR_SERV, DAT_FATURA_COBR_SERV) 

SELECT '161' AS COD_COBRANCA,COD_SERVICO,QTD_COBR_SERV,INS_COBR_SERV,SUM(VAL_COBR_SERV) AS TOTAL,VAL_PROMO_COBR_SERV, (SELECT GETDATE()-DAY(GETDATE())+1) AS DAT_FATURA_COBR_SERV 
FROM SIAF_PLUS.PRODUCAO.CLIENTES CLI
JOIN SIAF_PLUS.producao.COBRANCAS C ON cli.COD_CLIENTE = c.COD_CLIENTE
JOIN SIAF_PLUS.producao.COBRANCAS_SERVICOS CS ON C.COD_COBRANCA = CS.COD_COBRANCA	
WHERE CPF_CLIENTE = (SELECT CPF_CLIENTE FROM SIAF_PLUS.producao.CLIENTES WHERE COD_CLIENTE = 161) AND CLI.COD_CLIENTE <> 161
AND CS.DAT_FATURA_COBR_SERV < GETDATE()
GROUP BY COD_SERVICO,QTD_COBR_SERV,INS_COBR_SERV,VAL_PROMO_COBR_SERV


SELECT CLI.NOM_CLIENTE, CLI.APELIDO,  cs.DAT_FATURA_COBR_SERV,VAL_COBR_SERV
FROM SIAF_PLUS.PRODUCAO.CLIENTES CLI
JOIN SIAF_PLUS.producao.COBRANCAS C ON cli.COD_CLIENTE = c.COD_CLIENTE
JOIN SIAF_PLUS.producao.COBRANCAS_SERVICOS CS ON C.COD_COBRANCA = CS.COD_COBRANCA	
JOIN SIAF_PLUS.producao.SERVICOS S ON S.COD_SERVICO = CS.COD_SERVICO
WHERE CPF_CLIENTE = (SELECT CPF_CLIENTE FROM SIAF_PLUS.producao.CLIENTES WHERE COD_CLIENTE = 161) AND CLI.COD_CLIENTE <> 161
AND CS.DAT_FATURA_COBR_SERV <  GETDATE()

 SELECT SUM(VAL_cobr_serv) as total
 FROM SIAF_PLUS.PRODUCAO.CLIENTES CLI
 JOIN SIAF_PLUS.producao.COBRANCAS C ON cli.COD_CLIENTE = c.COD_CLIENTE
 JOIN SIAF_PLUS.producao.COBRANCAS_SERVICOS CS ON C.COD_COBRANCA = CS.COD_COBRANCA	
 JOIN SIAF_PLUS.producao.SERVICOS S ON S.COD_SERVICO = CS.COD_SERVICO
 WHERE CPF_CLIENTE = (SELECT CPF_CLIENTE FROM SIAF_PLUS.producao.CLIENTES WHERE COD_CLIENTE = 161) AND CLI.COD_CLIENTE <> 161
 AND CS.DAT_FATURA_COBR_SERV < GETDATE() 


select *  FROM SIAF_PLUS.producao.MENSALIDADES WHERE COD_COBRANCA = 161


select MES_PRESTACAO from SIAF_PLUS.producao.MENSALIDADES where COD_CLIENTE = 161


delete  from  siaf_plus.dbo.MENSALIDADES_SERVICOS where COD_MENSALIDADE = 878

select * from siaf_plus.dbo.MENSALIDADES_SERVICOS 
 

SELECT * FROM [SIAF_PLUS].producao.COBRANCAS
select * from SIAF_PLUS.producao.COBRANCAS where COD_CLIENTE = 161
select *  FROM SIAF_PLUS.producao.COBRANCAS_SERVICOS where COD_COBRANCA = 161

SELECT C.*, DES_SERVICO FROM SIAF_PLUS.producao.COBRANCAS_SERVICOS C, SIAF_PLUS.producao.SERVICOS S  WHERE C.COD_SERVICO=S.COD_SERVICO AND COD_COBRANCA= 161












