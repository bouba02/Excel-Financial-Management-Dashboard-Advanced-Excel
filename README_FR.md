# Dashboard Excel — Pilotage Financier Intégré | Excel Avancé

> **Système complet 6 onglets · SUMIFS · INDEX/MATCH · Alertes conditionnelles · Export comptable**  
> Ventes · Dépenses · Budget · Rentabilité · Comparaison N vs N-1 · Compatible Windows & Mac

![Dashboard Preview](Dashboard_Ngroup_HD.png)

🇬🇧 [English version available here](README.md)

---

## Problème Business

Les PME et micro-entreprises pilotent leur activité depuis des fichiers Excel
épars — aucune vision consolidée, clôtures manuelles chronophages, et aucune
alerte automatique sur les dérapages financiers.

**4 questions auxquelles le dashboard répond :**

| Question | Axe |
|---|---|
| Où en est ma trésorerie ce mois-ci ? | Temps réel |
| Mes dépenses dérapent-elles par rapport au budget ? | Contrôle |
| Quels clients génèrent le plus de CA ? | Commercial |
| Quelle est ma rentabilité vs l'année dernière ? | Comparatif |

---

## Architecture — 6 Onglets

```
┌─────────────────────────────────────────────────┐
│  ONGLET 1 : PARAMETRES                          │
│  Catégories dépenses · Règles comptables ·      │
│  Périodes reporting · Validation données         │
├─────────────────────────────────────────────────┤
│  ONGLET 2 : VENTES_REVENUS                      │
│  Date | Client | Montant | Catégorie | Statut   │
│  Détection doublons · Pivot clients              │
├─────────────────────────────────────────────────┤
│  ONGLET 3 : DEPENSES                            │
│  Date | Fournisseur | Montant | Catégorie        │
│  Validation montants · Alertes dépenses anorm.  │
├─────────────────────────────────────────────────┤
│  ONGLET 4 : TABLEAU_DE_BORD ⭐ Cœur du système  │
│  Zone 1 : KPIs (CA · Dépenses · Résultat · %)  │
│  Zone 2 : Graphiques (N vs N-1 · CA vs Dep.)   │
│  Zone 3 : Trimestriel · Répartition dépenses   │
│  Zone 4 : Top 10 clients · Détail mensuel       │
├─────────────────────────────────────────────────┤
│  ONGLET 5 : EXPORT_COMPTABLE                    │
│  Format standard expert-comptable               │
├─────────────────────────────────────────────────┤
│  ONGLET 6 : GUIDE_UTILISATION                   │
│  Mode d'emploi · FAQ · Dépannage                │
└─────────────────────────────────────────────────┘
```

---

## Formules Clés

```excel
// CA par client sur période
=SUMIFS(Ventes[Montant], Ventes[Client], [@Client],
        Ventes[Date], ">="&DateDebut,
        Ventes[Date], "<="&DateFin)

// Comparaison N vs N-1
=IFERROR(
    SUMIFS(Ventes[Montant], Ventes[Année], AnnéeRef)
    / SUMIFS(Ventes[Montant], Ventes[Année], AnnéeRef-1) - 1,
    0)

// Lookup dynamique catégorie
=INDEX(Parametres[Libellé],
       MATCH([@Catégorie], Parametres[Code], 0))
```

---

## Alertes Conditionnelles

```
🔴 Écart budget > 10%     → Surlignage rouge automatique
🟠 Dépense anormale       → Alerte vs historique
🟢 Tout sous contrôle     → Indicateur vert
```

---

## Dashboard — 4 Zones

![Dashboard](Dashboard_Ngroup_HD.png)

| Zone | Contenu |
|---|---|
| Zone 1 — KPIs | CA Total · Dépenses · Résultat Net · Marge % · N vs N-1 |
| Zone 2 — Graphiques | CA mensuel N vs N-1 (courbe) · CA vs Dépenses (barres empilées) |
| Zone 3 — Analyses | Performance trimestrielle T1→T4 · Répartition dépenses |
| Zone 4 — Détails | Top 10 clients (classement + % CA) · Tableau Jan–Déc |

---

## Stack Technique

| Outil | Usage |
|---|---|
| **Excel Avancé** | SUMIFS · INDEX/MATCH · Tableaux croisés dynamiques |
| **Mise en forme conditionnelle** | Alertes visuelles sur seuils métier |
| **Graphiques dynamiques** | Mise à jour automatique à chaque saisie |
| **Validation de données** | Listes déroulantes · Contrôles de saisie |

**Capacité :** 10 000+ lignes · Compatible Windows & Mac · Excel 2016+

---

## Structure du Repository

```
Excel-Dashboard-Pilotage-Financier/
├── README.md
├── README_FR.md
├── Dashboard_Ngroup_HD.png
└── Pilotage_Financier.xlsx
    ├── PARAMETRES
    ├── VENTES_REVENUS
    ├── DEPENSES
    ├── TABLEAU_DE_BORD
    ├── EXPORT_COMPTABLE
    └── GUIDE_UTILISATION
```

---

## Installation

```bash
git clone https://github.com/bouba02/Excel-Dashboard-Pilotage-Financier.git
```

Ouvrir `Pilotage_Financier.xlsx` dans Excel 2016+ ou Office 365.

---

## Auteur

**Boubacar Nikiema** — Data Analyst & Consultant BI

Spécialisé en dashboards financiers, Excel avancé et pilotage de la performance
pour PMEs avec Power BI, SQL, Python et Excel. Basé au Maroc, j'interviens
auprès d'entreprises en Afrique et en Europe francophone.

[![LinkedIn](https://img.shields.io/badge/LinkedIn-boubacar--nikiema-blue?logo=linkedin)](https://linkedin.com/in/boubacar-nikiema)
[![YouTube](https://img.shields.io/badge/YouTube-BoubacarDataAnalyst-red?logo=youtube)](https://youtube.com/@BoubacarDataAnalyst)
[![Email](https://img.shields.io/badge/Email-nikiemaboubacar%40gmail.com-gray?logo=gmail)](mailto:nikiemaboubacar@gmail.com)
[![Portfolio](https://img.shields.io/badge/Portfolio-data.ngroupmediadigital.com-green)](https://data.ngroupmediadigital.com)

---

*Code : MIT License · Compatible Excel 2016+ et Office 365*
