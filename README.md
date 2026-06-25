# check-diff

Tool CLI per confrontare due file Excel riga per riga e rilevare le differenze tra i dati.

---

## Prerequisiti

Prima di iniziare, assicurati di avere installato:

- **Python 3.11** — versione richiesta dal progetto
- **uv** — package manager Python usato per gestire l'ambiente virtuale e le dipendenze

### Installare Python 3.11

Scarica e installa da [python.org](https://www.python.org/downloads/) oppure tramite `pyenv`:

```bash
brew install pyenv
pyenv install 3.11
```

### Installare uv

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

Verifica l'installazione:

```bash
uv --version
```

---

## Setup su un nuovo computer

### 1. Clona il repository

```bash
git clone <url-del-repository>
cd check-diff
```

### 2. Crea l'ambiente virtuale e installa le dipendenze

```bash
uv sync
```

Questo comando crea automaticamente la cartella `.venv/` e installa tutte le dipendenze definite in `pyproject.toml`:

- `pandas` — lettura e manipolazione dei file Excel
- `openpyxl` — engine per i file `.xlsx`
- `tqdm` — barra di avanzamento durante il confronto

---

## Utilizzo

### 1. Posiziona i file Excel

Copia i due file da confrontare nella cartella `files/` con questi nomi esatti:

```
files/
├── old_file.xlsx   ← versione precedente
└── new_file.xlsx   ← versione nuova
```

### 2. Esegui il confronto

```bash
uv run basicCheck
```

### 3. Risultato

- L'output viene stampato a terminale con una barra di avanzamento per ogni foglio
- Il diff completo viene scritto nel file **`output.txt`** nella directory corrente

---

## Opzioni

| Flag | Alias | Default | Descrizione |
|------|-------|---------|-------------|
| `--precision` | `-p` | `6` | Numero di cifre decimali usate nel confronto dei valori numerici |

Esempio con precisione personalizzata:

```bash
uv run basicCheck --precision 2
```

---

## Cosa fa il tool

- Legge tutti i fogli presenti in entrambi i file `.xlsx`
- Segnala i fogli presenti solo nel file vecchio o solo nel file nuovo
- Per ogni foglio in comune, confronta le righe una a una
- Per le righe che differiscono, mostra colonna per colonna i valori `old` e `new`
- Normalizza i numeri (arrotondamento configurabile) per evitare falsi positivi da differenze di precisione floating point

### Output finale

- `RESULT: Files are identical.` — nessuna differenza trovata
- `RESULT: Files differ. See details above.` — differenze trovate, dettagli in `output.txt`

---

## Struttura del progetto

```
check-diff/
├── files/
│   ├── old_file.xlsx      # file da confrontare (non committato)
│   └── new_file.xlsx      # file da confrontare (non committato)
├── main.py                # logica principale
├── pyproject.toml         # configurazione progetto e dipendenze
├── uv.lock                # lockfile dipendenze
└── output.txt             # risultato dell'ultimo confronto (generato)
```
