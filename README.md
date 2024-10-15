# pun-fasce

Mostra i costi del PUN di un anno, diviso per mesi e fasce orarie (F1, F2 e F3).
I valori vengono scaricati dal sito [MercatoElettrico.org](https://www.mercatoelettrico.org/it-it/Home/Esiti/Elettricita/MGP/Statistiche/DatiStorici).

È stato aggiunto anche il valore della fascia mono-oraria e F23, quest'ultima basata su calcoli documentati non in modo perfetto ma che sembrano portare a risultati identici (o perlomeno simili) alle tariffe ufficiali.

## Installazione dei prerequisiti

`python3 -m pip install -r ./requirements.txt`

## Utilizzo

```bash
python3 ./pun-fasce.py <data_inizio> <data_fine> [-csv]
```
Argomenti:

* <data_inizio>: Data di inizio nel formato "YYYY-MM".
* <data_fine>: Data di fine nel formato "YYYY-MM".
* -csv (opzionale): Se fornito, esporta l'output in CSV.

Esempio:
`python3 ./pun-fasce.py "2024-01" "2024-09"`
```text
Mese	MO (€/kWh)	F1 (€/kWh)	F2 (€/kWh)	F3 (€/kWh)	F23 (€/kWh)
1/2024	0.099160	0.109650	0.105070	0.089060	0.096425
2/2024	0.087630	0.096150	0.094920	0.076810	0.085141
3/2024	0.088860	0.094930	0.094620	0.081320	0.087438
4/2024	0.086800	0.085570	0.101290	0.080540	0.090085
5/2024	0.094880	0.094660	0.111480	0.086240	0.097850
6/2024	0.103170	0.103810	0.116160	0.095430	0.104966
7/2024	0.112320	0.108670	0.130630	0.104770	0.116666
8/2024	0.128440	0.121670	0.147950	0.122190	0.134040
9/2024	0.117130	0.122330	0.131740	0.105630	0.117641
```
