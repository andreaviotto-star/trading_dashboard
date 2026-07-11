# 📊 Guida al Quant Portfolio Dashboard
### Versione 9.0 — Guida pratica per capire ogni sezione

---

## 🚀 Come avviare il dashboard

Apri il terminale nella cartella del progetto e scrivi:

```bash
# Installa le librerie (solo la prima volta)
pip install -r requirements.txt

# Avvia il dashboard
streamlit run app.py
```

Il browser si apre automaticamente su `http://localhost:8501`.
Per fermarlo: premi `Ctrl+C` nel terminale.

---

## ⚙️ Barra laterale sinistra (Impostazioni)

Questa barra controlla tutto il dashboard.

| Impostazione | Cosa fa |
|---|---|
| **Data directory** | La cartella dove si trovano i file `.xlsx` delle strategie |
| **Lookback (years)** | Quanti anni di storia mostrare nei grafici. Default: 10 anni |
| **Micro ($)** | Commissione per trade su contratti micro (MES, MNQ, MGC). Default: $1.50 |
| **Mini/Full ($)** | Commissione su contratti standard (ES, NQ, GC). Default: $15.00 |
| **Target daily risk ($)** | Usato dal Risk Parity: quanto rischio massimo giornaliero vuoi per sistema. Default: $500 |
| **📥 Generate Excel Report** | Crea e scarica un file Excel completo con tutti i dati |

> 💡 **Consiglio:** Cambia il lookback a 10 anni per vedere tutta la storia disponibile. Abbassalo a 2-3 anni per vedere solo la performance recente.

---

## 🖥 TAB 1 — Sistemi

**Cos'è:** Una tabella riassuntiva di tutte le 18 strategie, ognuna con i suoi numeri chiave.

### Colonne della tabella

| Colonna | Significato in parole semplici |
|---|---|
| **Src** | ✅ = dimensionamento letto dal ReadMe.txt correttamente. ⚠️ = valore di default usato (verifica) |
| **System** | Nome della strategia |
| **Sym** | Mercato: ES = S&P500, NQ = Nasdaq, GC = Oro, CL = Petrolio |
| **Contract** | Quanti contratti e di che tipo usa questa strategia |
| **# Trades** | Quante operazioni ha fatto in totale |
| **Net Profit** | Guadagno totale in dollari (commissioni incluse) |
| **Max DD** | **Massimo Drawdown** = la perdita peggiore subita dall'apice al minimo. Più è basso (meno negativo), meglio è |
| **Sharpe** | **Indice di Sharpe** = rapporto guadagno/rischio. Sopra 1 = buono, sopra 2 = ottimo, sotto 0 = problematico |
| **PF** | **Profit Factor** = somma guadagni ÷ somma perdite. Sopra 1.5 = buono, sopra 2 = ottimo |
| **Calmar** | Simile a Sharpe ma usa il drawdown invece della volatilità. Sopra 1 = accettabile |

### System Explorer (sezione sotto la tabella)

Seleziona una strategia dal menu per vedere:
- **I numeri dettagliati** (stessi indicatori sopra)
- **Grafico equity** = linea che sale = soldi guadagnati nel tempo. La zona rossa sotto = drawdown (perdite temporanee rispetto al picco)
- **Heatmap mensile** = griglia verde/rosso che mostra quale mese di ogni anno è stato positivo o negativo
- **Lista trade** = ogni singola operazione con data entrata, uscita, P&L

---

## 📦 TAB 2 — Portfolio

**Cos'è:** Simula il portafoglio combinando tutte le strategie insieme, come se le facessi girare tutte allo stesso tempo.

### ⚡ Simulatore What-If (in cima)

*"Cosa succede se raddoppio le dimensioni di tutto?"*

- **Moltiplicatore globale:** Trascina il cursore per applicare un moltiplicatore a tutti i contratti del ReadMe. Esempio: 2× significa il doppio dei contratti su ogni sistema.
- **La tabella a destra** mostra automaticamente come cambiano P&L, Drawdown, Sharpe e Calmar a 0.5×, 1×, 1.5×, 2×, 3×.
- **Slider individuali sotto:** Puoi anche regolare ogni sistema singolarmente.

> 💡 **Quando usarlo:** Prima di aumentare le dimensioni, guarda se il drawdown rimane accettabile. Se il Calmar scende sotto 1 raddoppiando i contratti, stai aumentando il rischio più del guadagno.

### Grafico portafoglio

Il grafico grande mostra:
- **Aree colorate impilate** = contributo di ogni sistema al totale
- **Linea nera/scura** = equity totale del portafoglio
- **Zona rossa in basso** = drawdown del portafoglio

### Contribuzione al P&L

Grafico a barre orizzontali: verde = sistema in profitto, rosso = sistema in perdita nel periodo selezionato.

### 🔍 Decomposizione Drawdown

*"Durante la peggior perdita del portafoglio, quale sistema era la causa principale?"*

- Mostra i 3 peggior episodi di drawdown del portafoglio
- Per ogni episodio: grafico a barre con la **% di colpa** di ogni sistema
- Esempio: "Durante il drawdown di -$12.000 da Gen a Giu 2022, il sistema GC Donchian era responsabile del 45%"

> 💡 **Quando è utile:** Se vedi che lo stesso sistema causa sempre i drawdown peggiori, valuta di ridurne le dimensioni o spegnerlo temporaneamente.

---

## 🔬 TAB 3 — Correlazione & Ottimizzazione

Questa tab ha tre sezioni principali.

### Heatmap di Correlazione (con Clustering)

**Cos'è la correlazione?** Misura quanto due strategie si muovono insieme.
- **+1.0** = si muovono esattamente allo stesso modo (molto rischioso averle entrambe — raddoppi il rischio senza diversificare)
- **0.0** = indipendenti (ideale)
- **-1.0** = si muovono in direzioni opposte (rarissimo nella pratica)

**Clustering:** Il grafico riordina automaticamente i sistemi raggruppando quelli simili. I sistemi vicini nella griglia tendono a comportarsi allo stesso modo.

**Rettangoli rossi:** Coppie di sistemi con correlazione sopra la soglia impostata (default 0.70). Se vedi molti rettangoli rossi, il portafoglio è concentrato — una perdita in un sistema probabilmente colpirà anche gli altri.

**Cluster Risk Score (0-100):** Punteggio di concentrazione del portafoglio. Sotto 30 = ben diversificato. Sopra 70 = troppo concentrato, considera di spegnere alcune strategie simili.

### ⚖️ Risk Parity Sizing

**Cos'è:** Un metodo matematico per bilanciare il rischio invece di bilanciare il capitale.

**Il problema del dimensionamento manuale:** GC (oro) è molto più volatile di ES (S&P500). Se metti 5 MGC e 5 MES, in realtà l'oro domina il rischio del portafoglio perché si muove di più.

**La soluzione Risk Parity:** Calcola automaticamente quanti contratti assegnare a ogni sistema in modo che ogni sistema contribuisca la stessa quantità di rischio giornaliero (il "Target daily risk $" nella barra laterale).

La tabella mostra:
- **Daily Vol 1ct** = quanto si muove in media al giorno un singolo contratto di quella strategia
- **Risk Parity N** = contratti suggeriti per raggiungere il target di rischio
- **Change** = ▲ aumenta contratti, ▼ diminuisce rispetto al ReadMe

Il grafico confronta l'equity con il dimensionamento del ReadMe (linea tratteggiata) vs Risk Parity (linea continua).

### 🎯 Portafogli Raccomandati

Tre portafogli pre-costruiti, ognuno con una logica diversa:

| Portfolio | Logica |
|---|---|
| 🏆 Max Sharpe | Prende le strategie con il miglior rapporto guadagno/rischio, escludendo quelle troppo correlate tra loro |
| 🌐 Max Diversificazione | Prende le strategie che si comportano in modo più diverso tra loro (correlazione minima) |
| 🛡 Min Drawdown | Prende le strategie che hanno subito le perdite peggiori più contenute |

Per ognuno viene mostrato: dimensionamento **Default** (dal ReadMe) vs **Ottimizzato** (calcolato automaticamente per massimizzare l'obiettivo).

> ⚠️ **Attenzione:** Questi portafogli usano solo un sottoinsieme dei 18 sistemi. Il P&L sarà più basso del portafoglio completo — l'obiettivo è un *rendimento corretto per il rischio* migliore, non un profitto assoluto più alto.

### 🔧 Ottimizzatore Personalizzato

Scegli tu quali sistemi includere e quale obiettivo ottimizzare (Sharpe, Calmar, Min DD, Max Return), poi clicca **Run Optimizer**.

L'algoritmo ("coordinate descent") prova diverse combinazioni di contratti per trovare il mix ottimale. Non è perfetto ma è un ottimo punto di partenza.

---

## 🔮 TAB 4 — Analisi Forward

Questa tab risponde alla domanda: *"Come stanno andando i sistemi adesso e cosa potrei aspettarmi in futuro?"*

### 🏥 Monitor di Salute dei Sistemi

Ogni sistema ha un semaforo basato sulle ultime 3 mesi (o 6 mesi) di performance:

| Semaforo | Significato | Cosa fare |
|---|---|---|
| 🟢 Verde | Sistema in buona forma, Sharpe alto, profit factor > 1.3 | Mantieni o aumenta leggermente |
| 🟡 Giallo | Performance nella media, attenzione | Monitora, non aumentare |
| 🔴 Rosso | Sistema in difficoltà, profit factor < 1, Sharpe negativo | Considera di ridurre o spegnere |

> 💡 **Importante:** Queste sono medie mobili, non profezie. Un sistema può essere rosso per 2 mesi e poi tornare verde. Usalo come campanello d'allarme, non come regola assoluta.

Il grafico "Rolling Sharpe nel tempo" mostra come l'Indice di Sharpe di ogni sistema è cambiato negli ultimi anni — utile per vedere se un sistema si sta deteriorando gradualmente.

### 🎲 Proiezione Monte Carlo

**Cos'è:** Il computer ripete 1000 volte una "simulazione del futuro" campionando casualmente i rendimenti storici giornalieri del portafoglio.

**Come leggere il grafico:**
- **Linea centrale (mediana)** = il risultato più probabile
- **Fascia chiara (25-75%)** = metà delle simulazioni cadono in questa zona
- **Fascia molto chiara (5-95%)** = la grande maggioranza delle simulazioni
- **Asse X** = giorni di trading nel futuro (63 = ~3 mesi, 126 = ~6 mesi)

**La tabella "Probabilità di Drawdown"** mostra: se continuo con le dimensioni attuali, qual è la probabilità di perdere almeno $X nei prossimi mesi?

> ⚠️ **Disclaimer:** Il Monte Carlo si basa sul passato. Non prevede crisi improvvise, cambiamenti di regime o eventi straordinari. È uno strumento di gestione del rischio, non una garanzia.

### 📋 Raccomandazioni di Sizing

Tabella riassuntiva che combina il semaforo di salute con una raccomandazione pratica per ogni sistema:
- 🟢 Sharpe alto + PF alto → "Considera di aumentare a N contratti"
- 🟡 Nella media → "Mantieni, monitora"
- 🔴 In difficoltà → "Considera di ridurre a N o mettere in pausa"

### 🌡️ Analisi del Regime di Mercato

**Cos'è un regime di mercato?** Il mercato non si comporta sempre allo stesso modo. Ci sono periodi di bassa volatilità (mercato tranquillo) e periodi di alta volatilità (mercato turbolento).

Questa sezione classifica ogni giorno in uno di tre regimi:

| Regime | Cosa significa |
|---|---|
| 😴 **Low Vol** (Bassa volatilità) | Mercati tranquilli, trend stabili. Le strategie trend-following funzionano bene |
| 😐 **Medium Vol** (Media) | Condizioni normali |
| 🔥 **High Vol** (Alta volatilità) | Mercati frenetici, molti falsi breakout. Le strategie mean-reverting tendono a soffrire |

**Indicatore attuale:** In alto vedi in quale regime siamo *adesso*. Utile per sapere quali sistemi dovrebbero essere in buona forma in questo momento.

**Tabella performance per regime:** Per ogni sistema vedi il P&L medio, win rate e profit factor separatamente in ognuno dei tre regimi. Se un sistema guadagna bene solo in Low Vol e perde in High Vol, sai che in periodi di turbolenza è meglio ridurlo.

---

## 📥 Export Excel

Clicca **📥 Generate Excel Report** nella barra laterale. Scarica un file con 6 fogli:

1. **Summary** — tabella riassuntiva di tutti i sistemi con KPI
2. **Detailed KPIs** — metriche complete
3. **Trade Log** — ogni singolo trade di ogni sistema
4. **Correlation Matrix** — matrice di correlazione con colori
5. **Risk Parity Sizing** — analisi del dimensionamento risk parity
6. **Optimizer Results** — risultati dell'ultimo ottimizzatore eseguito

---

## 📖 Glossario rapido

| Termine | Spiegazione semplice |
|---|---|
| **Equity Curve** | La "linea dei soldi" — mostra l'andamento cumulativo del P&L nel tempo |
| **Drawdown** | Perdita rispetto al massimo precedente. Se ero a +$10.000 e ora sono a +$7.000, il drawdown è -$3.000 |
| **Max Drawdown** | La perdita peggiore mai subita dall'apice al minimo |
| **Sharpe Ratio** | Rendimento annuo ÷ volatilità annua. Il "voto" di efficienza della strategia |
| **Profit Factor** | Totale guadagni ÷ totale perdite. 1.5 significa che per ogni $1 perso ne guadagni $1.50 |
| **Calmar Ratio** | Rendimento annuo ÷ max drawdown. Misura quanto guadagni rispetto al rischio peggiore |
| **Correlazione (ρ)** | Da -1 a +1: quanto due cose si muovono insieme. 0 = indipendenti. +1 = identici |
| **Risk Parity** | Tecnica di sizing: dai più contratti alle strategie "tranquille" e meno a quelle "violente" |
| **Regime** | Stato del mercato: tranquillo (low vol), normale (medium), frenetico (high vol) |
| **Rolling** | "Mobile": calcolato sull'ultima finestra temporale, aggiornato ogni giorno |
| **MES/MNQ/MGC** | Contratti micro: MES = mini S&P500 micro, MNQ = Nasdaq micro, MGC = Oro micro |
| **ES/NQ/GC/CL** | Contratti standard: ES = S&P500 mini, NQ = Nasdaq mini, GC = Oro, CL = Petrolio |

---

## ❓ Domande frequenti

**Il grafico mostra solo 2 anni anche se ho 10 anni di dati — perché?**
Controlla il cursore "Lookback (years)" nella barra laterale. Mettilo a 10.

**Il P&L mostrato è reale o in backtest?**
È il P&L storico calcolato dai file Excel esportati da TradeStation. È backtest, non live trading. I numeri includono le commissioni impostate nella barra laterale.

**Il sistema X è 🔴 rosso — devo spegnerlo subito?**
No. Il semaforo è basato sulle ultime settimane. Un sistema può essere temporaneamente in difficoltà e poi recuperare. Usalo come segnale di attenzione, poi valuta se ci sono cause fondamentali (cambio di regime, over-ottimizzazione).

**La correlazione tra ES-long e ES-short è -0.8 — è normale?**
Sì, è normalissimo. Un sistema che va long sull'ES e uno che va short sullo stesso mercato tendono ad avere correlazione negativa. Il portafoglio beneficia di averli entrambi.

**Il Risk Parity suggerisce di aumentare molto i contratti di ES — è sicuro?**
Il Risk Parity bilancia il *rischio*, non il capitale. ES è molto meno volatile di GC, quindi il modello suggerisce più contratti. Verifica sempre che il drawdown del portafoglio risk parity sia accettabile prima di applicare i cambiamenti.

---

*Guida v9.0 — Quant Portfolio Dashboard*
