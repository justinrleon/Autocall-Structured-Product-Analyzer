# Structured Products Analyzer — Phoenix Autocall

**Projet personnel de finance quantitative | Excel · VBA · Monte-Carlo · Produits structurés**

Ce projet est un outil pédagogique permettant de **configurer, simuler et analyser un Phoenix Autocall sur un seul sous-jacent**. Il combine des données historiques, des scénarios déterministes et un moteur Monte-Carlo développé en VBA pour illustrer les mécanismes de coupon conditionnel, de rappel anticipé et de protection conditionnelle du capital.

> **À retenir :** ce projet est une démonstration quantitative, **pas** une proposition commerciale, un conseil en investissement ou un outil de cotation bancaire. Les résultats simulés ne garantissent aucun rendement futur.

![Aperçu de l'interface de paramétrage](docs/apercu_parametres.png)

*Capture d'une configuration illustrative du classeur ; les paramètres et estimations peuvent changer selon le sous-jacent et la période retenus.*

## Démarrage rapide

1. Télécharger et ouvrir [`Analyse_produit_structuré_autocall.xlsm`](Analyse_produit_structur%C3%A9_autocall.xlsm) **dans Microsoft Excel de bureau**. Le fichier contient déjà son code VBA : **aucun import de `.bas` n'est nécessaire**.
2. Dans `SOUS_JACENTS`, choisir un actif en **`E5`** ; dans `PARAMETRES`, personnaliser le produit et vérifier le mode de calibration en **`E21`**.
3. Dans `SIMULATION`, fixer le nombre de tirages en **`P5`** (par exemple `1000` pour un premier test), puis lancer **`LancerMonteCarloPhoenix`** avec le bouton prévu ou **`Alt + F8`**.
4. Lire les courbes dans `SIMULATION`, les statistiques de risque dans `COMPARAISON` et la synthèse dans `FICHE_CLIENT`.

➡️ **[Guide d'utilisation détaillé](#guide-dutilisation-détaillé)** — choix des paramètres, modes de trajectoire, fonctionnement des macros et dépannage.


## Fonctionnalités

- **8 sous-jacents disponibles**, sélectionnables un par un : CAC 40 (`^FCHI`), EURO STOXX 50 (`^STOXX50E`), S&P 500 (`^GSPC`), NASDAQ-100 (`^NDX`), SPY (`SPY`), QQQ (`QQQ`), LVMH (`MC.PA`) et TotalEnergies (`TTE.PA`).
- **Paramétrage du Phoenix** : nominal, maturité, fréquence d'observation, coupon annuel, barrières de coupon / rappel / protection et première observation de rappel.
- **Deux modes de trajectoire** dans `TRAJECTOIRES` : `Historique` (cours passés) et `Personnalisée` (variations saisies manuellement).
- **Monte-Carlo en VBA** : de **100 à 20 000 trajectoires**, simulées au pas mensuel ; application des conditions du Phoenix aux dates d'observation.
- **Visualisation de la distribution du sous-jacent** : minimum, 1er quartile (Q1), moyenne, 3e quartile (Q3), maximum à chaque observation.
- **Analyse des résultats financiers** : coupons, capital remboursé, gain net, durée de vie effective, fréquence des rappels et des pertes en capital, quantiles et scénarios extrêmes.
- **Tableaux de bord** dans `SIMULATION`, `COMPARAISON`, `FICHE_CLIENT` et `ACCUEIL`, reliés aux résultats du moteur.

## Exemple de Phoenix étudié

| Paramètre | Configuration illustrative |
| --- | ---: |
| Nominal | 10 000 € |
| Maturité maximale | 6 ans |
| Observations | Trimestrielles (4/an) |
| Coupon conditionnel | 8 % par an, soit 2 % par trimestre éligible |
| Barrière de coupon | 70 % du niveau initial |
| Barrière d'Autocall | 100 % du niveau initial |
| Première observation de rappel | Observation n° 4 (après un an) |
| Barrière de protection finale | 60 % du niveau initial |
| Mémoire des coupons | Non |

**Fonctionnement :** à chaque date d'observation, un coupon est versé si le sous-jacent atteint au moins la barrière de coupon. À partir de la première observation éligible, si le sous-jacent atteint la barrière de rappel, le nominal est remboursé et le produit prend fin. En l'absence de rappel, la barrière de protection du capital est évaluée **uniquement à l'échéance** : au-dessus ou à la barrière, le nominal est intégralement remboursé ; en dessous, le remboursement suit la performance finale du sous-jacent.

Ces conditions sont **modélisées pour le projet** et ne décrivent pas une émission réelle.

## Méthodologie quantitative

### 1. Calibration statistique

Les prix mensuels sont utilisés pour estimer les log-rendements :

```text
r_t = ln(P_t / P_(t-1))
mu_log_annuel = moyenne(r_t) × 12
sigma_annuel = ecart_type(r_t) × sqrt(12)
```

Deux modes de calibration sont disponibles :

- **Historique** : dérive logarithmique et volatilité estimées sur les rendements mensuels disponibles des mois complets.
- **Hypothèses manuelles** : dérive logarithmique annualisée et volatilité annualisée saisies par l'utilisateur, utiles pour des scénarios de sensibilité.

La **volatilité historique** est donc estimée automatiquement dans le mode Historique ; elle ne correspond pas à la volatilité implicite d'options.

### 2. Simulation du sous-jacent

Le moteur VBA produit des trajectoires sous une dynamique lognormale, au pas mensuel :

```text
S_(m+1) = S_m × exp(mu_log_mensuel + sigma_mensuel × Z_m)
Z_m ~ N(0,1)
```

Les tirages normaux sont générés à partir de nombres pseudo-aléatoires uniformes (transformation de Box-Muller). La calibration est **statistique**, et non risque-neutre. Le modèle ne fournit donc **pas un prix théorique négociable** du Phoenix.

Pour tracer des statistiques à horizon commun, l'évolution du sous-jacent peut continuer à être simulée après un éventuel rappel ; **les flux du Phoenix s'arrêtent, eux, au rappel**.

### 3. Flux du Phoenix

Pour une observation éligible :

```text
Coupon = Nominal × taux_coupon_annuel / frequence
         si niveau >= barriere_coupon ; sinon 0

Si date de rappel autorisee et niveau >= barriere_rappel :
    remboursement = Nominal ; fin du produit

Sinon, à maturité :
    remboursement = Nominal                  si niveau_final >= barriere_protection
    remboursement = Nominal × niveau_final  sinon (niveau relatif, base 1)

Gain net nominal = coupons cumules + capital rembourse - nominal
```

Les coupons et remboursements sont exprimés en **montants nominaux non actualisés**. Le rendement total présenté n'est pas un TRI annualisé.

### 4. Résultats et risques

Le moteur produit la distribution des gains et pertes et calcule notamment :

- le gain et les coupons moyens ;
- les probabilités *simulées* de rappel, de perte en capital et de gain net négatif ;
- la durée de vie effective moyenne ;
- le minimum / maximum observés parmi les tirages ;
- les quantiles de gain à 5 % et 95 % et la moyenne des tirages les plus défavorables.

Les cinq courbes du graphique (minimum, Q1, moyenne, Q3, maximum) sont des **statistiques transversales à chaque date**, et non cinq trajectoires individuelles continues.

## Organisation du classeur

| Feuille | Fonction |
| --- | --- |
| `ACCUEIL` | Navigation et indicateurs de synthèse |
| `SOUS_JACENTS` | Choix de l'actif analysé |
| `PARAMETRES` | Conditions du produit, hypothèses et estimations |
| `TRAJECTOIRES` | Scénario personnalisé **ou** historique, avec flux détaillés |
| `SIMULATION` | Nombre de tirages, courbes MC et indicateurs synthétiques |
| `COMPARAISON` | Comparaison scénario choisi / distribution Monte-Carlo |
| `FICHE_CLIENT` | Synthèse vulgarisée du Phoenix et de ses risques |
| `MC_MOTEUR` | Niveaux des trajectoires générées par VBA |
| `MC_RESULTATS` | Résultats financiers de chaque simulation |
| `DONNEES_*`, `DONNEES` | Historique des huit actifs et données de référence |

## Guide d'utilisation détaillé

Ce guide décrit le classeur **`Analyse_produit_structuré_autocall.xlsm`** joint à ce dépôt. Les adresses ci-dessous ont été vérifiées dans cette version. Le fichier partagé contient des résultats Monte-Carlo déjà calculés ; **il faut relancer la simulation après tout changement d'hypothèse**.

### 0. Ouvrir le fichier et activer les fonctionnalités

- Utiliser **Microsoft Excel de bureau** avec prise en charge des macros VBA ; **Excel en ligne / l'aperçu GitHub ne peut pas exécuter le code VBA**.
- Ouvrir le classeur `.xlsm`. Si Excel bloque ses macros, suivre les règles de sécurité locales et ne les autoriser **que si le fichier et son origine sont fiables**.
- La présence du fichier `xl/vbaProject.bin` dans le classeur confirme que le projet VBA est intégré. On peut le consulter dans Excel avec **`Alt + F11`**, dossier **Modules**. **Ne pas réimporter les anciens modules `.bas`**, au risque de créer des procédures en double.
- Les formes d'`ACCUEIL` servent de raccourcis lorsqu'une macro leur est affectée. En cas de bouton absent ou inactif, ouvrir **`Alt + F8`** et choisir la procédure directement. Il n'est pas nécessaire de relancer une macro d'installation de boutons pour utiliser les simulations.

### 1. Choisir le sous-jacent

Ouvrir **`SOUS_JACENTS`** et utiliser la liste déroulante de **`E5`**. Un seul sous-jacent est analysé à la fois : CAC 40, EURO STOXX 50, S&P 500, NASDAQ-100, SPY, QQQ, LVMH ou TotalEnergies.

Les informations d'identification (code, devise, marché, ticker Yahoo Finance et nombre de cours) sont affichées dans cette même feuille. Les historiques sont embarqués dans les onglets `DONNEES_*` ; **aucune connexion Internet n'est requise pour simuler**.

### 2. Configurer le produit dans `PARAMETRES`

La feuille distingue les **entrées modifiables** des **valeurs estimées/calculées** :

| Cellule | Paramètre | Comment l'utiliser |
| --- | --- | --- |
| `E7` | Nominal (€) | Montant investi, par exemple `10000`. |
| `E8` | Niveau initial | **Base 100 fixe**, ne pas modifier. |
| `E9` | Maturité maximale | Entre **1 et 10 ans**. |
| `E10` | Observations par an | `1` (annuel), `2` (semestriel), `4` (trimestriel), `12` (mensuel). |
| `E11` | Coupon annuel conditionnel | Par exemple **`8 %`**, versés à proportion de la fréquence si la condition est satisfaite. |
| `E12` | Barrière de coupon | Par exemple **`70 %`** du niveau initial. |
| `E13` | Barrière de rappel | Par exemple **`100 %`** du niveau initial. |
| `E14` | Première observation de rappel | **Numéro d'observation**, pas numéro de mois : avec `E10=4`, la valeur `4` autorise le premier rappel après **un an**. |
| `E15` | Barrière de protection finale | Par exemple **`60 %`**, testée uniquement à maturité en l'absence de rappel. |
| `E17` | Dérive annuelle — hypothèse | Utilisée **uniquement** en mode manuel ; vérifier la convention de dérive dans le code VBA (voir note technique). |
| `E18` | Volatilité annuelle — hypothèse | Utilisée **uniquement** en mode manuel, par exemple **`22 %`**. |
| `E20` | Début du backtest | Date de lancement d'une trajectoire **historique** ; **ne fixe pas la date de départ du Monte-Carlo**. |
| `E21` | Calibration MC | Choisir exactement **`Historique`** ou **`Hypothèses manuelles`** dans la liste. |

Les cellules **`E24:E31`** contiennent notamment le nombre de cours, la dérive et la volatilité historiques, le nombre d'observations et les paramètres retenus pour la simulation. Elles sont destinées à la lecture : éviter d'écraser leurs formules. **`E19` reflète le nombre de tirages sélectionné dans `SIMULATION!P5`.**

**Choisir le mode de calibration :**

- **`Historique`** (`PARAMETRES!E21`) : le moteur estime la moyenne et la volatilité des rendements logarithmiques mensuels de l'actif sélectionné à partir des mois complets. Les hypothèses `E17` et `E18` n'interviennent pas.
- **`Hypothèses manuelles`** : le moteur exploite les entrées `E17` et `E18`, utiles pour étudier l'effet d'une hausse de volatilité ou d'une autre hypothèse de tendance. Ce mode reste une **simulation statistique**, pas une cotation de marché.

> **Convention de dérive à vérifier avant une étude quantitative formelle :** dans l'interface, `E17` est nommé « Dérive LOG annuelle ». Le moteur VBA développé lors du projet inclut toutefois une conversion de dérive avec un terme `−½σ²` en mode manuel. Il faut harmoniser l'intitulé et cette conversion dans le code si l'on souhaite interpréter `E17` strictement comme une dérive logarithmique. Le mode `Historique` utilise directement la moyenne des log-rendements mensuels estimés.

### 3. Explorer `TRAJECTOIRES` : deux modes distincts

Cette feuille reproduit les **flux d'un seul Phoenix sur une seule trajectoire** ; elle ne remplace pas les milliers de scénarios Monte-Carlo.

**Mode `Historique`**

1. Choisir **`Historique`** dans **`TRAJECTOIRES!E5`**.
2. Sélectionner une date passée dans **`PARAMETRES!E20`**, par exemple janvier 2020.
3. Examiner **`TRAJECTOIRES!E7`** : `OK` signifie que la trajectoire historique requise est disponible ; `INCOMPLET` indique que la période ne permet pas de calculer tout le scénario dans les conditions sélectionnées.
4. Lire les niveaux historiques normalisés, les coupons, le rappel éventuel, le capital remboursé et le gain dans le tableau et la zone de synthèse. Les résultats sont **hypothétiques** : ils appliquent les règles choisies à des cours passés.

**Mode `Personnalisée`**

1. Choisir **`Personnalisée`** dans **`TRAJECTOIRES!E5`**.
2. Saisir les variations entre deux observations **dans `D14:D133` uniquement** pour les périodes actives (par exemple `-10 %`, `+5 %`, `0 %`). Les premières lignes correspondent à l'observation n° 1, n° 2, etc. : **la fréquence est déterminée par `PARAMETRES!E10`**.
3. La colonne **`F`** calcule le niveau en base 100 : `F14 = 100 × (1 + D14)` puis `F15 = F14 × (1 + D15)`, et ainsi de suite.
4. Lire les colonnes **`G` à `L`** : coupon, indicateur d'Autocall, état actif, remboursement, flux et coupons cumulés. Le produit ne verse plus de flux après rappel.

**Exemple :** si les quatre premières variations trimestrielles sont `−10 %`, `−20 %`, `−10 %`, puis `+10 %`, les niveaux passent de 100 à **90**, **72**, **64,8**, puis **71,28**. Au quatrième trimestre, le sous-jacent est au-dessus de la barrière de coupon à 70, mais sous la barrière de rappel à 100 : le coupon est éligible et le produit n'est pas rappelé.

Les deux modes de `TRAJECTOIRES` sont **indépendants** du mode de calibration du Monte-Carlo : on peut donc étudier un scénario historique tout en simulant avec des hypothèses manuelles.

### 4. Lancer le Monte-Carlo VBA

1. Ouvrir **`SIMULATION`** et saisir le nombre de trajectoires dans **`P5`**. Le moteur du projet accepte **de 100 à 20 000** tirages ; pour un premier essai, utiliser `1000`.
2. Vérifier le sous-jacent et les paramètres avant le lancement.
3. Exécuter le bouton associé à **`LancerMonteCarloPhoenix`**. À défaut, utiliser **`Alt + F8` → `LancerMonteCarloPhoenix` → Exécuter**.
4. Attendre le message de fin de traitement. La cellule **`SIMULATION!U5`** affiche également un état d'avancement ou de fin.
5. Contrôler que `MC_MOTEUR` contient les niveaux simulés et que **`MC_RESULTATS`** contient les gains, coupons, remboursements et durées. Ce sont des **feuilles techniques** : ne pas modifier manuellement leurs valeurs.

Le calcul simule le sous-jacent **chaque mois**, vérifie les conditions contractuelles **aux dates d'observation sélectionnées**, et poursuit la trajectoire de marché sur l'horizon total pour pouvoir construire les statistiques graphiques. **Les flux du Phoenix s'arrêtent lors de son Autocall.**

### 5. Lire `SIMULATION`, `COMPARAISON` et `FICHE_CLIENT`

| Feuille | Ce qu'il faut regarder | Interprétation |
| --- | --- | --- |
| `SIMULATION` | Nombre de tirages, gain/coupons/durée moyens, risque de perte et de rappel, cinq courbes **min / Q1 / moyenne / Q3 / max** | Les cinq courbes résument les **niveaux du sous-jacent à chaque date**, pas cinq Phoenix différents. |
| `MC_RESULTATS` | Une ligne par tirage : coupons, capital, gain net, rendement total, durée, rappel, perte en capital, indice à la sortie | Données détaillées, utiles pour les vérifications. |
| `COMPARAISON` | Colonnes **Scénario choisi / MC pire tirage / MC moyenne / MC meilleur tirage**, et statistiques à 5 % et 95 % | Le « pire » et le « meilleur » sont définis **par le gain net** parmi les simulations. |
| `FICHE_CLIENT` | Caractéristiques du contrat, résultat du scénario choisi, gain MC moyen, pire gain simulé et probabilité de perte | Présentation pédagogique vulgarisée, pas fiche commerciale officielle. |
| `ACCUEIL` | Accès aux feuilles et chiffres clés | Navigation et résumé de la configuration courante. |

Le gain est calculé en montants **nominaux, non actualisés** : `coupons cumulés + capital remboursé − nominal investi`. La « probabilité de perte en capital » désigne la proportion de tirages où le **capital remboursé** est inférieur au nominal, même si des coupons ont été perçus. Le rendement total n'est pas un TRI annualisé.

### 6. Mettre les tableaux à jour après chaque changement

**Après une modification du sous-jacent, du produit ou des hypothèses, relancer `LancerMonteCarloPhoenix`.** Les anciennes simulations restent visibles jusqu'au prochain lancement et ne décrivent plus forcément la configuration affichée.

La liaison des résultats financiers avec `ACCUEIL`, `SIMULATION`, `COMPARAISON` et `FICHE_CLIENT` repose sur les formules installées par la procédure **`RelierTableauxPhoenix`**. Dans la version documentée, cette macro est destinée à être appelée à la fin du Monte-Carlo ; en cas de tableaux non actualisés (ou si la plage change lorsque l'on passe de 1 000 à 10 000 tirages), on peut l'exécuter séparément via **`Alt + F8`**. Il n'est **pas nécessaire de l'importer de nouveau**.

### 7. Dépannage rapide

| Problème | Vérification / solution |
| --- | --- |
| Les boutons ne réagissent pas | Vérifier l'autorisation des macros ; essayer `Alt + F8` pour exécuter directement la procédure. Les boutons de navigation sont des formes Excel affectées à une `Public Sub`. |
| `MC_RESULTATS` est vide | Lancer **`LancerMonteCarloPhoenix`** ; vérifier `SIMULATION!P5`, la sélection de l'actif et la disponibilité des cours. |
| `FICHE_CLIENT` ou `COMPARAISON` n'affiche pas les nouveaux chiffres | Si `MC_RESULTATS` est rempli, exécuter **`RelierTableauxPhoenix`** puis recalculer avec **`Ctrl + Alt + F9`**. |
| Le mode historique affiche `INCOMPLET` | Vérifier `PARAMETRES!E20` et la maturité. Les historiques ne contiennent pas toujours une période suffisante jusqu'à la date visée. |
| Le mode personnalisé ne réagit pas | Vérifier **`TRAJECTOIRES!E5 = Personnalisée`** et saisir les variations dans **`D14:D133`**, non dans les cellules de calcul de la colonne F. |
| Résultats incohérents après modification des paramètres | Relancer le moteur MC et s'assurer que le mode `PARAMETRES!E21` correspond aux hypothèses souhaitées. |
| Le code VBA disparaît à l'enregistrement | Enregistrer le fichier sous le format **`.xlsm`** et non `.xlsx`. |

### 8. Deux parcours de démonstration

**Démonstration « compréhension du produit » (sans VBA) :** choisir LVMH → conserver nominal 10 000 €, coupon 8 % et observations trimestrielles → choisir `TRAJECTOIRES!E5 = Personnalisée` → saisir plusieurs baisses/reprises en colonne D → expliquer la présence ou l'absence des coupons et le déclenchement éventuel du rappel.

**Démonstration « risque quantitatif » (avec VBA) :** choisir un sous-jacent → conserver `PARAMETRES!E21 = Historique` → lancer 1 000 ou 10 000 simulations → examiner les courbes de `SIMULATION` → comparer le pire tirage, le gain moyen et la probabilité de perte dans `COMPARAISON` → conclure avec `FICHE_CLIENT`, en rappelant les limites statistiques du modèle.

## Sources des données

Les historiques de cours mensuels ont été constitués à partir de données Yahoo Finance via `yfinance` et intégrés **en valeurs statiques** au classeur (période de travail : **2015–2026**). Les données ne sont **pas mises à jour automatiquement**. Les cours utilisés sont des cours de clôture non ajustés ; cela limite notamment l'interprétation de certains rendements pour les actions et ETF.

## Limites et avertissements

- Modèle lognormal à volatilité constante et paramètres calibrés sur le passé : il ne capture pas intégralement les sauts, queues épaisses ni changements de régime.
- Absence de calibration aux volatilités implicites, de courbe de taux, de modélisation complète des dividendes et de valorisation risque-neutre.
- Risque de crédit / défaut de l'émetteur, liquidité, frais et fiscalité **non modélisés** dans les résultats chiffrés.
- Hypothèses contractuelles simplifiées (notamment **sans mémoire des coupons**).
- Montants de gain/perte non actualisés ; extrêmes simulés **non garantis** ; probabilités **conditionnelles au modèle**.
- Le tirage pseudo-aléatoire n'est pas fixé par une graine reproductible dans la version documentée : deux exécutions peuvent différer.

## Publier le projet sur GitHub

Structure suggérée :

```text
structured-products-analyzer/
├── README.md
├── Analyse_produit_structuré_autocall.xlsm
└── docs/
    ├── Presentation_Phoenix_Autocall.pdf
    └── apercu_parametres.png
```

Le code VBA est **déjà inclus dans le fichier `.xlsm`** fourni ici. Pour le rendre facilement consultable sur GitHub sans devoir ouvrir Excel, tu peux aussi **exporter tes modules définitifs en fichiers `.bas`** depuis l'éditeur VBA (`Fichier → Exporter un fichier`), puis les ranger dans `src/`. Vérifie qu'ils correspondent bien à la version finale du classeur.

Évite d'ajouter des fichiers temporaires Excel (`~$...`), des identifiants personnels ou des données financières non autorisées à la diffusion.

## Document de présentation

Voir [la fiche projet PDF](docs/Presentation_Phoenix_Autocall.pdf) pour une présentation synthétique destinée à un recruteur ou à un entretien.

---

*Projet indépendant à vocation pédagogique. Aucune affiliation ni validation par un émetteur, un courtier ou un fournisseur de données n'est revendiquée.*
