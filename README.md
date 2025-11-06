# JIRA Worklog Riport Készítő

Python alkalmazás JIRA worklogok lekérdezésére és riport készítésére macOS GUI-val.

## Funkciók

- GUI felület macOS-re (tkinter)
- JIRA worklogok lekérdezése Personal Access Token használatával
- Felhasználó specifikus worklog keresés
- JQL alapú szűrés
- Excel riport generálás

## Telepítés

1. Klónozd le a repository-t
2. Hozz létre virtual environment-et:
```bash
python3 -m venv venv
source venv/bin/activate
```

3. Telepítsd a függőségeket:
```bash
pip install -r requirements.txt
```

4. Hozd létre az `auth.json` fájlt a következő tartalommal:
```json
{
  "jira": {
    "url": "https://jira.teszt.hu",
    "pat": "YOUR_PERSONAL_ACCESS_TOKEN"
  }
}
```

## Használat

```bash
python jira_worklog_app.py
```

1. Add meg a JIRA felhasználónevet (pl.: kasnyikl)
2. Add meg a JQL lekérdezést (pl.: project = MYPROJECT)
3. Kattints a "Lekérdezés indítása" gombra
4. A riport automatikusan elkészül a `reports` mappában

## Megjegyzés

Az `auth.json` fájl .gitignore-ban van, ne commitold a verziókezelőbe!
