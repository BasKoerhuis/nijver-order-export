# Nijver Order Export Converter

🚀 **Live demo:** https://nijver-order-export.vercel.app

Webapp die Nijver order-export CSV's omzet naar opgemaakte Excel-bestanden in
de Nijver-huisstijl. Aggregeert measures per complex op basis van NL-SfB-codes.

## Lokaal draaien

```bash
npm install
npm run dev
```

Open http://localhost:5173

## Hoe het werkt

Upload twee CSV's:

- **`project.csv`** — order-metadata (auteur, ordercode, corporatie)
- **`complexes.csv`** — één rij per complex met alle measures en prijzen

De app:

1. Parseert beide CSV's met PapaParse
2. Groepeert measures per complex op NL-SfB-hoofdstukcode → categorieën:
   Dak, Gevel, Installatie, Interieur, Overig, Vloer
3. Genereert een xlsx met twee bladen (Uitgangspunten + Prijs) met volledige
   cel-styling via `xlsx-js-style`

## Tech stack

- React 18 + Vite
- xlsx-js-style (cel-styling support)
- papaparse (CSV parsing)
- lucide-react (iconen)

## Deploy

Push naar GitHub, importeer op [vercel.com/new](https://vercel.com/new) — Vercel
detecteert Vite automatisch.

---

Gebouwd voor [Nijver](https://nijver.nl).
