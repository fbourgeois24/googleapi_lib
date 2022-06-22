# googleapi_lib

## Google sheet
### Créer l'objet
```python
sheet = googleapi_lib.google_sheet("1TMtA3_FatJ3Gifh_cQTGtDPdI", "keys.json")
```

### Lire une plage
```python
data = sheet.read("A1:B10")
```