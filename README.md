# FareHub

Pesquisador local de tarifas aéreas em classe Economy. Para cada par de aeroportos,
consulta saídas em 7, 30 e 60 dias ou em uma data de embarque escolhida e mostra
a menor tarifa retornada por companhia em cada data. Aceita pesquisa individual
ou planilha XLSX com até 10 trechos. O menu separa **Trecho específico**,
**Embarque específico**, **Lista em Excel** e **Consumo das APIs**.
Antes de cada consulta, é possível deixar a seleção automática ou escolher
uma chave específica, inclusive a **API 1 (teste)**.

## Iniciar

```powershell
python -m pip install -r requirements.txt
python pesquisa_tarfia.py
```

Abra <http://127.0.0.1:8502>. A planilha de exemplo está no botão **Baixar exemplo
de Excel**. Use códigos IATA de três letras nas colunas `origem` e `destino`.
Na pesquisa em lote, as colunas opcionais `taxa latam`, `taxa gol` e `taxa azul`
recebem a taxa de embarque em BRL para cada companhia naquele trecho. Coluna
ausente ou célula vazia vale R$ 0,00. O XLSX exportado inclui **Taxa (BRL)** e
**Tarifa net (BRL)**, calculada como total retornado pela API menos essa taxa.
Esses valores também aparecem nos resultados da página. O **share total** é a
tarifa total da menor opção de uma companhia dividida pela soma das tarifas
totais das companhias exibidas no mesmo trecho e dia. O **share net** usa o
mesmo cálculo com as tarifas net. Os percentuais refletem as opções da pesquisa,
não participação de mercado. O arquivo exportado inclui a aba inicial **Planilha Simples**
(tabela consolidada com todas as opções encontradas, contendo ADVP, Data da viagem,
Origem, Destino, Cia, Número do voo, Hora do voo, Trecho, Tarifa net, Taxa e Tarifa total),
além das abas detalhadas por dia com agrupamento por trecho e shares, e a aba
**Comparativo Cia vs Cia**.

O sistema também inclui o painel **Comparativo Cia vs Cia**, calculando a diferença
percentual direta entre companhias que operam o mesmo trecho na mesma data
(ex: LATAM 15,2% menor que GOL, LATAM 8,5% maior que Azul), tanto para tarifas Net
quanto para tarifas Totais. Essa análise também é gerada em aba dedicada no Excel
(`Comparativo Cia vs Cia`), com formatação condicional e percentuais com sinal (+/-).

## Fontes e cotas

O aplicativo lê todas as seções do `.env` existente, inclusive chaves repetidas.
Numera as chaves na ordem do arquivo. Usa primeiro as contas SerpApi, começando
pela conta `KARENALIXAVI@GMAIL.COM`; tenta RapidAPI quando elas falham ou acabam.
As credenciais nunca são enviadas ao navegador nem gravadas no banco local.

As cotas mensais configuradas são 250 requisições por chave SerpApi e 150 por
chave RapidAPI. O consumo SerpApi é consultado pela Account API gratuita. A RapidAPI normalmente
informa saldo e limite nos cabeçalhos de uma resposta; antes da primeira chamada,
ou se o serviço não devolver esses cabeçalhos, o saldo fica indisponível. O botão
**Consultar saldo RapidAPI** faz uma busca de teste em cada chave RapidAPI para
obter esses cabeçalhos e pode consumir uma requisição por chave. O painel
também conta todas as tentativas feitas por este aplicativo e o consumo estimado.
O banco `farehub.sqlite3` guarda esse histórico e os resultados para exportação.

O fallback presume o serviço RapidAPI `google-flights2.p.rapidapi.com`. Se as
chaves pertencerem a outro serviço, configure `RAPIDAPI_HOST` e
`RAPIDAPI_SEARCH_PATH` no ambiente antes de iniciar. O `APPLICATION_ID_RAPIDAPI`
do `.env` não identifica o serviço contratado.

Preços são valores totais para uma pessoa, só ida, em BRL. A SerpApi e a
RapidAPI DataCrawler retornaram `price` como inteiro nas buscas testadas. Os
detalhes de reserva da RapidAPI trouxeram preços por vendedor, mas não tarifa
base, impostos ou desconto separados. Por isso a tela mostra a **Tarifa Total**
retornada pela API. As taxas da planilha são valores informados pelo usuário,
separados de qualquer detalhamento da API. O valor exato e seus componentes
precisam ser confirmados no checkout da companhia.
