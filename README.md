# Excel Notes

## Como somar vendas por mês de cada produto

Pagina principal > Relatorios > Vendas > por mes > Analitico

Como copiar datas nas linhas vazias

```
 A  B
 1  =A1
    =SE(A2<>""; A2; SE(E(A2=""; A1<>""); A1; B1))
    ...
 2
```

Apos calcular datas, copie tudo e "colar especial" (Ctrl+Shift+V) os textos, números e datas

As datas são valores em texto. Para converter: =DATA.VALOR(L2)

Como eliminar duplicatas

```
A  B
1  =A1
   =SE(A2=A1;0;1)
2  ...
```

colar especial e classificar

### Criar variaveis temporarios

`=CONCATENAR(codigo, mes)`

Matriz de codigos com meses

```
       MES
CODIGO 1       2       3
137263 1372631 1372632 ...
...
```

com a identificacao de codigo e mes, calcular vendas por mes:

``` 
=SOMASE($coluna_codigo_e_mes na planilha de vendas;
        PLANILHA_MATRIZ codigo_e_mes;
        $coluna_quantidade por venda)
```

Estoque se não existe é zero:

```
=SE(É.NÃO.DISP(PROCV(CODIGO;$COLUNAS_COD_E_QTDE;2;0));0;
               PROCV(CODIGO;$COLUNAS_COD_E_QTDE;2;0))
```

## Aggregate values that match then apply function

`{=MÉDIA(SE(A7:A13=C7;B7:B13))}` then Ctrl+Shift+Enter

```
COD VAL 
A   5   B
A   5
B   10
A   5
A   5
B   30
A   10
```

## Procedimento para calcular total de vendas por cliente e associar com a vendedora

1. Relatorios > Vendas por numero > Sintetico

Selecionar datas desejadas

2. Relatorios > Clientes > Lista de clientes

Preencher Vendedor = `PONTUAL`, `CHEN`
Preencher Secundario = `CELIA`, `MARCIA`, etc

Não imprimir enderecos nem observacoes

3. Classificar vendas por nome de cliente

4. Usar formula `ESQUERDA(nome; 20)` porque nomes muito longos sao separados em 2 linhas

5. Formula para associar vendedora com valores de vendas:

```
=SE(É.NÃO.DISP(PROCV(NOME_CLIENTE;$LISTA_DE_CLIENTES_COM_VENDEDORA;2;0));"PONTUAL";PROCV(NOME_CLIENTE;$LISTA_DE_CLIENTES_COM_VENDEDORA;2;0))
```

`PONTUAL` é usado quando o valor desejado não é encontrado

6. Formula para calcular total

```
=SOMASE($Coluna_Grande_do_nome_das_vendedoras;Nome_da_vendedora;$Coluna_dos_valores)
```
