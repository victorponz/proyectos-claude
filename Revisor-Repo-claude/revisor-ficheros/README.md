# Revisor automàtic d'entregues DWES (Symfony)

Revisa entregues d'alumnes Symfony segons la rúbrica de RA i genera un Excel
amb una fila per alumne (estat, evidència i suggeriment per cada RA).

## 1. Instal·lació
```
pip install -r requirements.txt
```

## 2. Configuració

**Rúbrica:** edita `rubrica.md` amb el text exacte dels 9 RA i els seus
criteris d'avaluació (els tens als documents Word que ja vas generar per
al mapeig RA ↔ projectes).

**Clau de l'API:**
```
export ANTHROPIC_API_KEY="la-teua-clau"
```

**Entregues:** organitza-les amb una subcarpeta per alumne:
```
entregas/
  Alumne1/   (repo Symfony complet: src/, templates/, config/...)
  Alumne2/
  ...
```

## 3. Ús
```
python revisar_entregas.py --entregas ./entregas --rubrica rubrica.md --salida informe_revisio.xlsx
```

## Notes
- S'exclouen automàticament `vendor/`, `var/`, `node_modules/`, `.git/`.
- Si un projecte és molt gran, només s'envien els primers ~180.000 caràcters,
  prioritzant `src/`, `templates/`, `config/` i `migrations/`.
- Cada error de revisió (alumne concret) queda registrat a la columna
  "Error" de l'Excel, sense aturar la resta del procés.
- Pensat com a suport a la correcció, no com a nota automàtica: revisa
  l'evidència de cada RA abans de posar la qualificació final.
