# The Galleon · Plancia di comando

Cruscotto gestionale per The Galleon Pub & Pizzeria: rotta verso l'utile, food cost, costo del lavoro, conto economico, menu engineering e controllo sprechi.

**Demo online:** apri `index.html` (pubblicato con GitHub Pages). **I numeri sono DI ESEMPIO, inventati.**

## Come funziona
- `sistema/pipeline.py` gira ogni notte **sul PC del locale**: legge chiusure di cassa, fatture elettroniche e presenze, calcola tutto e scrive `report/app/dati.js` + l'app.
- **Conta settimanale**: `conta.html` si apre dal telefono, si contano i 9 ingredienti di classe A (5 minuti) e si invia il file `conta_AAAA-MM-GG.csv` in `dati/inbox/conte`. Nella vista *Sprechi* compare lo scarto della settimana ingrediente per ingrediente.
- L'app (`index.html`) legge solo `dati.js` (numeri già sintetizzati) e funziona anche offline con doppio clic.
- Dettagli in [`sistema/ARCHITETTURA.md`](sistema/ARCHITETTURA.md).

> ⚠️ Questo repository è pubblico: qui dentro vanno solo codice e dati di esempio. I dati reali del locale restano sul PC (sono esclusi da `.gitignore`).
