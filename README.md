# Prenota con Elia – setup

Struttura: `index.html` (sito statico, va su GitHub Pages) → Apps Script (Calendar + Sheet + mail) → n8n (alarm).

## 1. Backend (Apps Script)
1. Crea un Google Sheet "Prenotazioni Elia". Estensioni → Apps Script.
2. Incolla `apps-script/Code.gs`. In Impostazioni progetto attiva "Mostra appsscript.json" e incolla `apps-script/appsscript.json`.
3. In `CONFIG` imposta `CALENDAR_ID` (calendario di Elia, condiviso con l'account che pubblica lo script con permesso di modifica) e gli `ORARI` reali.
4. Impostazioni progetto → Proprietà script:
   - `SITE_URL` = `https://TUOUTENTE.github.io/NOME-REPO/` (con lo slash finale)
   - `N8N_WEBHOOK_URL` = URL del webhook n8n (facoltativa)
   - `N8N_SECRET` = stringa a tua scelta (facoltativa)
5. Esegui `setup()` una volta e autorizza: crea il foglio e il trigger dei promemoria (ogni giorno alle 18:00, mail ai clienti con appuntamento il giorno dopo).
6. Distribuisci → Nuova distribuzione → App web → Esegui come: *me*, Chi ha accesso: *Chiunque*. Copia l'URL `/exec`.
   Ogni modifica a `Code.gs` richiede una nuova versione della distribuzione.

## 2. Sito (GitHub)
1. In `index.html` sostituisci `CONFIG.API` con l'URL `/exec`.
2. Aggiungi una `privacy-policy.html` (il link è già nel form: trattandosi di dati legati a salute/benessere serve un'informativa vera).
3. Carica i file su GitHub e attiva Pages (Settings → Pages → branch main).

## 3. n8n (alarm a ogni prenotazione)
Webhook (POST) → controlla l'header `X-Secret` → Switch su `type`:
- `booking`: `nome, cognome, telefono, email, data, ora` → manda la notifica a Elia (Telegram, push, WhatsApp…)
- `swap`: oggetti `a` e `b` con vecchi contatti e nuovo orario di entrambi
- `cancel`: dati della prenotazione annullata

I promemoria del giorno prima li invia già Apps Script. Se preferisci farli in n8n, elimina il trigger `sendReminders` e leggi il foglio (stato = attivo, data = domani, promemoria vuoto).

## Come funziona la privacy
- Nel calendario ogni evento ha come titolo solo il nome.
- Cognome, telefono, email e token stanno nel foglio (il token è salvato come hash).
- Lo scambio mostra agli altri clienti solo data e ora, mai nomi.
- Il token personale arriva nel link della mail di conferma e resta nel browser; chi lo perde non può più gestire l'appuntamento da solo.
- Regole: una prenotazione futura per email; niente scambi o annullamenti a meno di 24 ore (`MIN_SWAP_H`).

## Da verificare prima di andare online
Non ho potuto provare il flusso su un tuo Calendar reale. Fai un giro di prova: prenota, scambia tra due email diverse, annulla, e controlla titolo evento, righe del foglio, mail e webhook n8n.
