# Revisor automàtic d'entregues DWES (GitHub Classroom)

Llista els repos d'una organització de GitHub Classroom amb un prefix
concret, els clona temporalment un a un, els avalua amb Claude segons la
rúbrica de RA, i genera un Excel amb una fila per alumne (estat, evidència
i suggeriment per cada RA). Cap repo queda guardat al disc: es clona en
una carpeta temporal i s'esborra automàticament en acabar cadascun.

## 1. Instal·lació
```
pip install -r requirements.txt
```
Cal tindre `git` instal·lat (ja sol estar-hi a la majoria de sistemes).

## 2. Configuració

**Rúbrica:** edita `rubrica.md` amb el text exacte dels 9 RA i els seus
criteris d'avaluació (els tens als documents Word que ja vas generar per
al mapeig RA ↔ projectes).

**Tokens:**

```
export ANTHROPIC_API_KEY="la-teua-clau-anthropic"
export GITHUB_TOKEN="ghp_la-teua-clau-github"
```
El token de GitHub necessita permís de lectura sobre els repositoris de
l'organització (scope `repo` si els repos de classroom són privats).

**Roster (opcional):** si vols que a l'Excel aparega el nom real de
l'alumne en lloc del seu usuari de GitHub, crea un CSV amb columnes
`github_username,nom_alumne` (és el format que exporta GitHub Classroom
des de "Manage students that joined with no roster identity" o el roster
del propi Classroom) i passa'l amb `--roster`.

## 3. Ús
```
python revisar_entregas.py \
  --org la-meua-organitzacio \
  --prefix whatsapp-clone \
  --rubrica rubrica.md \
  --roster roster.csv \
  --salida informe_revisio.xlsx
```

`--prefix` és el prefix que GitHub Classroom posa a tots els repos d'eixe
assignment (per exemple, si els repos es diuen
`whatsapp-clone-jgarcia23`, `whatsapp-clone-mlopez45`..., el prefix és
`whatsapp-clone`).

## Notes
- S'exclouen automàticament `vendor/`, `var/`, `node_modules/`, `.git/`.
- Si un projecte és molt gran, només s'envien els primers ~180.000
  caràcters, prioritzant `src/`, `templates/`, `config/` i `migrations/`.
- Cada error (de clonatge o d'avaluació) queda registrat a la columna
  "Error" de l'Excel per a eixe alumne, sense aturar la resta del procés.
- Pensat com a suport a la correcció, no com a nota automàtica: revisa
  l'evidència de cada RA abans de posar la qualificació final.
