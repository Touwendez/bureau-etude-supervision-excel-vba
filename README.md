# Bureau d’étude — Supervision sous Excel/VBA (Master SMaRT)

Projet académique de supervision : IHM (UserForms), événements Excel, timers non bloquants, pilotage par Grafcet, et supervision d’un procédé de fabrication avec arrêt d’urgence et historisation.

---

## Fonctionnalités
- IHM de commande Marche/Arrêt avec gestion d’état, verrouillage et retour visuel.
- Acquisition/gestion de données : génération, affichage graphique, calcul moyenne/écart-type, gestion d’événements Excel.
- Timers non bloquants avec `Application.OnTime` (clignoteur).
- Grafcet sous VBA : états, transitions, affichage couleur, mode manuel/automatique.
- Supervision d’un procédé : séquences de fabrication (fraisage/perçage/changement pièce), indicateurs, arrêt d’urgence avec reprise “temps restant”, sauvegarde d’historique (save.xls).

---

## Architecture (vue globale)

```mermaid
flowchart LR
  A[UserForms / IHM] --> B[Modules VBA]
  B --> C[Feuilles Excel (données, affichage, graphes)]
  B --> D[Timers Application.OnTime]
  B --> E[Grafcet: états/transitions]
  E --> F[Process de fabrication simulé]
  F --> G[Historisation vers save.xls (DDE)]



---

## Fichiers Excel (à télécharger)
- [tp1_exo2.xlsm](src/tp1_exo2.xlsm)
- [TP2 exo 2.xlsm](src/TP2%20exo%202.xlsm)
- [TP3.xlsm](src/TP3.xlsm)
- [mini projet VBA.xlsm](src/mini%20projet%20VBA.xlsm)
- [tp4_supervision.xlsx](src/tp4_supervision.xlsx)

---

## Aperçu visuel
### IHM Marche / Arrêt (UserForm)
![img1](assets/img1.png)

### Statistiques + événement Worksheet_Change
![img2](assets/img2.png)

### Mesure / stats temps de réaction
![img3](assets/img3.png)

### IHM Form_Exo1 (état vert / rouge)
![img4](assets/img4.png)

### Grafcet (mode manuel / automatique)
![img5](assets/img5.png)
