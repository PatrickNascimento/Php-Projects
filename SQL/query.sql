Helpers de SQL


AGREGAÇÕES
1. Ordenar com um tipo

Frequentemente, todos os seus dados precisam de ordem. A cláusula SQL “ORDER BY” organiza os dados em ordem alfabética ou numérica. Consequentemente, os valores semelhantes classificam-se juntos no que parece ser mais um grupo. No entanto, os grupos aparentes são um resultado de uma ordenação, não são verdadeiros grupos. ORDER BY exibe cada registro, enquanto que um grupo pode representar vários registros.
2. Reduzir valores semelhantes em um grupo

A maior diferença entre a classificação e agrupamento é esta: dados classificados exibem todos os registros (dentro dos limites de qualquer critério de limitação) e dados agrupados, não. A cláusula GROUP BY reduz valores semelhantes em um único registro. Por exemplo, uma cláusula GROUP BY pode retornar uma lista única de códigos postais de uma fonte que repete os valores:
Listagem 1: Exemplo de agrupamento
```
SELECT CEP
FROM Clientes
GROUP BY CEP
```
Inclua apenas as colunas que definem o grupo tanto no GROUP BY quanto no SELECT das listas de colunas. Em outras palavras, a lista resultante de SELECT deve coincidir com a lista do GROUP BY, com uma exceção: A lista do SELECT pode incluir funções de agregação (GROUP BY não permite funções agregadas.).
Tenha em mente que GROUP BY não irá classificar os grupos resultantes. Para organizar os grupos em ordem alfabética ou numérica, adicione uma cláusula ORDER BY (# 1). Além disso, você não pode se referir a um campo, aliás na cláusula GROUP BY. Colunas do grupo devem estar nos dados subjacentes, mas eles não devem aparecer nos resultados.
3. Limitar os dados antes de serem agrupados

Você pode limitar os dados dos grupos GROUP BY, adicionando uma cláusula WHERE. Por exemplo, a instrução a seguir retorna uma lista única de códigos postais para os clientes apenas em São Paulo:
Listagem 2: Filtrando antes de agrupar
```
SELECT CEP
FROM Clientes
WHERE Estado = 'SP'
GROUP BY ZIP
```
É importante lembrar que os filtros de dados do WHERE antes da cláusula GROUP BY são quem avaliam os dados.
Como GROUP BY, WHERE não suporta funções de agregação.
4. Retornar todos os grupos

Quando você usa WHERE para filtrar dados, os grupos resultantes exibem apenas os registros que você especificar. Os dados que se encaixam na definição do grupo, mas não atendem às condições da cláusula não farão parte do grupo. Inclua ALL quando você quiser incluir todos os dados, independentemente da condição WHERE. Por exemplo, a adição de ALL para a instrução anterior retorna todos os grupos CEP, não apenas aqueles em São Paulo:
Listagem 3: Adição da instrução ALL
```
SELECT CEP
FROM Clientes
WHERE Estado = 'SP'
GROUP BY ALL CEP
```
Neste caso, as duas cláusulas estão em conflito e você provavelmente não iria usar ALL neste caminho. ALL vem a calhar quando você usa um agregado para avaliar uma coluna. Por exemplo, a seguinte instrução conta o número de clientes em cada código postal de São Paulo, ao mesmo tempo, exibindo valores postais dos outros:
Listagem 4: Contando registros
```
SELECT CEP, Count(CEP) AS ContClientesPorCEP
FROM Clientes
WHERE Estado = 'SP'
GROUP BY ALL CEP
```
Os grupos resultantes compreendem todos os valores postais nos dados subjacentes. No entanto, a coluna agregada (ContClientesPorCEP) iria mostrar 0 para qualquer grupo que não seja um código postal de São Paulo válido.
Consultas remotas não suportam GROUP BY ALL.
5. Limitar os dados depois de agrupados

A cláusula WHERE (# 3) avalia os dados antes de a cláusula GROUP BY o fazer. Quando você quiser limitar os dados depois que agrupados, use HAVING. Muitas vezes, o resultado será o mesmo se você usar WHERE ou HAVING, mas é importante lembrar que as cláusulas não são intercambiáveis. Aqui está uma boa orientação a seguir quando você estiver em dúvida: use WHERE para filtrar os registros, use HAVING para filtrar grupos.
Normalmente, você vai usar HAVING para avaliar um grupo usando um agregado. Por exemplo, a instrução a seguir retorna uma lista única de códigos postais, mas a lista pode não incluir todos os CEPs na fonte de dados subjacente:
Listagem 5: Utilizando a cláusula HAVING
```
SELECT CEP, Count(CEP) AS ClientesPorCEP
FROM Clientes
GROUP BY CEP
HAVING Count(CEP) = 1
```
Apenas os grupos com apenas um cliente sairão no resultado.
3. Dar uma boa olhada no WHERE e HAVING

Se você ainda está confuso sobre onde e quando usar HAVING, aplique as seguintes diretrizes:
- WHERE vem antes de GROUP BY; SQL avalia a cláusula WHERE antes de seus grupos de registros.
- HAVING vem depois de GROUP BY; SQL avalia o HAVING após seus grupos de registros.
7. Resumir valores agrupados com agregados

O agrupamento de dados pode ajudar a analisar os dados, mas às vezes você vai precisar de um pouco mais informações do que apenas os próprios grupos. Você pode adicionar uma função de agregação para resumir dados agrupados. Por exemplo, a declaração a seguir exibe um subtotal para cada ordem:
Listagem 6: Usando funções de agregação
```
SELECT IDVenda, Sum(Custo * Quantidade) AS TotalVendido
FROM ItensVendidos
GROUP BY IDVenda
```
Tal como acontece com qualquer outro grupo, as listas dos SELECT e GROUP BY devem corresponder umas às outras. Incluir um agregado na cláusula SELECT é a única exceção a esta regra.
8. Resumir o conjunto

Você ainda pode resumir os dados exibindo um subtotal para cada grupo. O operador SQL “ROLLUP” exibe um registro extra, um subtotal, para cada grupo. Esse registro é o resultado da avaliação de todos os registros dentro de cada grupo usando uma função agregada. A declaração a seguir totaliza a coluna OrderTotal para cada grupo:
Listagem 7: Usando o operador ROLLUP
```
SELECT Cliente, NumeroVenda, Sum(Custo * Quantidade) AS TotalVendido
FROM ItensVendidos
GROUP BY Clientes, NumeroVenda
WITH ROLLUP
```
A linha ROLLUP para um grupo com dois valores TotalVendido de 20 e 25 teriam de apresentar um TotalVendido = 45. O primeiro registro em um resultado ROLLUP é único porque ele avalia todos os registros do grupo. Esse valor é um total geral para todo o conjunto de registros.
ROLLUP não suporta DISTINCT em funções agregadas ou a cláusula GROUP BY ALL.
9. Resumir cada coluna

O operador CUBE vai um passo além do que ROLLUP, retornando totais para cada valor em cada grupo. Os resultados são semelhantes aos ROLLUP, mas CUBE inclui um registo adicional para cada coluna no grupo. A declaração a seguir exibe um subtotal para cada grupo e um total adicional para cada cliente:
Listagem 8: Utilizando o operador CUBE
```
SELECT Clientes, NumeroVenda, Sum(Custo * Quantidade) AS TotalVendido
FROM ItensVendidos
GROUP BY Clientes, NumeroVenda
WITH CUBE
```
CUBE dá o resumo mais abrangente. Ele não só faz o trabalho de ambos os agregados e ROLLUP, mas também avalia as outras colunas que definem o grupo. Em outras palavras, CUBE resume todas as combinações possíveis de coluna.
CUBE não suporta GROUP BY ALL.
10. Traga a ordenação para os resumos

Quando os resultados de um CUBE são confusos (e geralmente são), adicione a função de agrupamento da seguinte forma:
Listagem 9: Função de agrupamento junto com CUBE
```
SELECT GROUPING(Customer), OrderNumber, Sum(Cost * Quantity) AS OrderTotal
 FROM Orders
 GROUP BY Customer, OrderNumber
 WITH CUBE
```
Os resultados incluem dois valores adicionais para cada linha:
- O valor 1 indica que o valor para a esquerda é um valor resumo - o resultado do operador ROLLUP ou CUBE.
- O valor 0 indica que o valor para a esquerda é um registro de detalhe produzido pela cláusula GROUP BY original.

REGISTROS DUPLICADOS

```
    select  count(CPF_CLIENTE) AS TOTALREG
    FROM [SIAF_PLUS].producao.CLIENTES 
    GROUP BY  CPF_CLIENTE
    HAVING Count(*) > 1
```

#####################################################################################################


BEGIN TRASACTION


```
BEGIN TRANSACTION
...
     IF @@ERROR = 0
      COMMIT
      ELSE
      ROLLBACK

#############################  SAMPLE

BEGIN TRANSACTION              
IF exists (SELECT * FROM [SIAF_PLUS].producao.COBRANCAS WHERE COD_CLIENTE = 161) 
  select 1+1 
    COMMIT; 
ELSE 
  select 2+2 
    COMMIT;
```
     


QUERY BUSCAR NOME DAS COLUNAS (COLUMNS)

```
SELECT COLUMN_NAME
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'Employees'
ORDER BY ORDINAL_POSITION
```
      

pegando nome servidor

```
select
   server_id,
   name,
   product,
   provider,  -- OLE DB provider
   data_source,
   is_remote_login_enabled,
   modify_date
from
   sys.servers
```


CRIANDO SCRIPT DE INSERT COM SELECT

```
select 'insert into Table values(Id=' + Id + ', name=' + name + ')' from Users
```


CLONAR UM REGISTRO

insert into ORDENS
SELECT * from ORDENS where COD_ORDEM = 10464;



SQL DATA TYPES


Dinheiro ->   format (12345678.90, 'c', 'pt-br') as MONEY


PEGANDO O PRIMEIRO DIA DO MES
--SQl-SERVER

SELECT GETDATE()-DAY(GETDATE())+1 AS FIRST_DAY_OF_DATE





