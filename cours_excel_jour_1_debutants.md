---
layout: default
title: Cours Excel -- Jour 1
permalink: /jour_1/
---
![Infographie pédagogique – Excel Jour 1](jour1_banner.jpeg)
# Prise en main, calculs, mise en forme et graphiques

**Formation :** 1ère année -- Soft Skills  
**Établissement :** Institut Supérieur DakarTech  
**Module :** Excel -- 21 heures au total  
**Public :** Étudiants débutants  
**Durée :** Première journée de formation

---

## Présentation du cours

Bienvenue dans ce premier cours consacré à Microsoft Excel.

L'objectif n'est pas d'apprendre une longue liste de commandes par cœur. Vous allez apprendre à utiliser Excel en réalisant progressivement des tâches concrètes : saisir des informations, effectuer des calculs, présenter proprement un tableau et transformer des chiffres en graphiques.

Nous allons avancer étape par étape. Chaque nouvelle notion sera expliquée, illustrée par un exemple, puis mise en pratique directement dans Excel.

**Important :** le programme de cette première journée porte sur les fonctions usuelles, les calculs, les formats de cellule, la création de graphiques et les fonctions complexes. Les macros et l'initiation à VBA sont prévues plus tard dans le module.

## 1. Objectifs pédagogiques

À la fin de cette séance, vous devrez être capable de :

- Vous repérer dans un classeur et une feuille de calcul.
- Saisir et modifier des données.
- Identifier une cellule et une plage de cellules.
- Écrire une formule simple.
- Utiliser les fonctions `SOMME`, `MOYENNE`, `MAX` et `MIN`.
- Mettre en forme un tableau.
- Appliquer des formats numériques, monétaires, de date et de pourcentage.
- Créer un graphique simple et le personnaliser.
- Utiliser les fonctions `SI`, `ET`, `OU` et `SOMME.SI`.

## 2. Découverte d'Excel

### 2.1. À quoi sert Excel ?

Microsoft Excel est un logiciel tableur. Il permet notamment de :

- Organiser des informations dans des tableaux.
- Effectuer des calculs automatiquement.
- Analyser des données.
- Présenter des résultats sous forme de graphiques.

**Exemples d'utilisation :**

- Suivre les notes d'une classe.
- Établir un budget.
- Enregistrer les paiements des étudiants.
- Calculer les ventes d'une entreprise.
- Suivre les dépenses d'un service.

### 2.2. Vocabulaire essentiel

| Terme | Définition |
|---|---|
| Classeur | Fichier Excel pouvant contenir plusieurs feuilles. |
| Feuille | Espace de travail composé de lignes et de colonnes. |
| Colonne | Ensemble vertical de cellules, identifié par une lettre : A, B, C... |
| Ligne | Ensemble horizontal de cellules, identifié par un numéro : 1, 2, 3... |
| Cellule | Intersection d'une ligne et d'une colonne. |
| Référence | Adresse d'une cellule, par exemple `B3`. |
| Plage | Ensemble de cellules, par exemple `B2:B8`. |

**Exemple :** `C4` désigne la cellule située dans la colonne C et la ligne 4.

La plage `B2:B6` comprend les cellules B2, B3, B4, B5 et B6.

### 2.3. Première manipulation

1. Ouvrez Microsoft Excel.
2. Créez un nouveau classeur vierge.
3. Repérez les lettres des colonnes et les numéros des lignes.
4. Cliquez sur la cellule A1.
5. Saisissez `Nom`, puis validez.
6. Cliquez sur B1 et saisissez `Note`.
7. Enregistrez le fichier sous le nom `Cours_Excel_Jour1.xlsx`.

**À retenir :** une cellule est toujours repérée par la lettre de sa colonne suivie du numéro de sa ligne.

## 3. Saisir et organiser des données

Avant de faire des calculs, il faut disposer de données bien organisées.

Saisissez le tableau suivant dans Excel à partir de la cellule A1 :

| Étudiant | Informatique | Mathématiques | Anglais |
|---|---:|---:|---:|
| Fall | 15 | 12 | 10 |
| Diop | 8 | 14 | 11 |
| Ba | 17 | 16 | 13 |
| Sarr | 9 | 10 | 12 |
| Ndiaye | 14 | 13 | 16 |

### Quelques bonnes habitudes

- Placez un titre clair au-dessus de chaque colonne.
- Saisissez une information par cellule.
- Ne laissez pas de ligne vide au milieu du tableau.
- Évitez de fusionner les cellules dans la zone de données.
- Utilisez le même type de données dans une même colonne.

**Question :** dans ce tableau, que représente une ligne ? Et que représente une colonne ?

## 4. Comprendre les formules

Une formule permet de demander à Excel d'effectuer un calcul.

### 4.1. La règle principale

Toute formule commence par le signe égal `=`.

Exemples :

```excel
=5+3
=10-4
=6*2
=20/5
```

Les opérateurs les plus utilisés sont :

| Opérateur | Opération | Exemple |
|---|---|---|
| `+` | Addition | `=A1+B1` |
| `-` | Soustraction | `=A1-B1` |
| `*` | Multiplication | `=A1*B1` |
| `/` | Division | `=A1/B1` |
| `^` | Puissance | `=A1^2` |

### 4.2. Calculer avec les références des cellules

Supposons que la cellule B2 contienne `15` et la cellule C2 contienne `12`.

Dans D2, saisissez :

```excel
=B2+C2
```

Excel additionne les valeurs de B2 et C2. Si vous modifiez ensuite B2 ou C2, le résultat se met à jour automatiquement.

**Pourquoi utiliser les références ?**

Parce qu'une formule liée aux cellules est réutilisable. Il est préférable d'écrire `=B2+C2` plutôt que de saisir directement `=15+12`.

### 4.3. Recopier une formule

Lorsqu'une formule doit être appliquée à plusieurs lignes :

1. Saisissez la formule dans la première cellule.
2. Sélectionnez cette cellule.
3. Placez le pointeur sur le petit carré en bas à droite de la sélection.
4. Faites glisser vers le bas.

Excel adapte normalement les références pour chaque ligne.

**Exemple :** si D2 contient `=B2+C2`, la recopie vers D3 produit `=B3+C3`.

## 5. Les fonctions usuelles

Une fonction est une formule déjà préparée par Excel. Elle permet d'effectuer un calcul sans devoir construire toute l'opération soi-même.

### 5.1. SOMME

La fonction `SOMME` additionne plusieurs valeurs.

```excel
=SOMME(B2:B6)
```

Cette formule additionne les valeurs comprises entre B2 et B6.

### 5.2. MOYENNE

La fonction `MOYENNE` calcule la moyenne des valeurs numériques.

```excel
=MOYENNE(B2:B6)
```

### 5.3. MAX et MIN

- `MAX` renvoie la valeur la plus élevée.
- `MIN` renvoie la valeur la plus faible.

```excel
=MAX(B2:B6)
=MIN(B2:B6)
```

### 5.4. NB et NBVAL

- `NB` compte les cellules qui contiennent des nombres.
- `NBVAL` compte les cellules non vides.

```excel
=NB(B2:B6)
=NBVAL(A2:A6)
```

### Application : compléter le tableau des étudiants

Reprenez le tableau de la section 3 et ajoutez trois colonnes : **Total**, **Moyenne** et **Décision**.

En supposant que les notes de Fall sont en B2, C2 et D2 :

Dans E2, calculez le total :

```excel
=SOMME(B2:D2)
```

Dans F2, calculez la moyenne :

```excel
=MOYENNE(B2:D2)
```

Recopiez les deux formules vers le bas.

Pour obtenir la meilleure moyenne de la classe :

```excel
=MAX(F2:F6)
```

Pour obtenir la moyenne la plus faible :

```excel
=MIN(F2:F6)
```

**Vérification :** les moyennes attendues sont :

- Fall : 12,33
- Diop : 11
- Ba : 15,33
- Sarr : 10,33
- Ndiaye : 14,33

## 6. Les formats de cellule

Le format de cellule permet de contrôler la manière dont une valeur apparaît dans Excel.

### 6.1. Les formats courants

| Format | Exemple | Utilisation |
|---|---|---|
| Standard | `1250` | Affichage général |
| Nombre | `1 250,00` | Valeurs numériques |
| Monétaire | `1 250 FCFA` | Montants financiers |
| Pourcentage | `25 %` | Taux et proportions |
| Date | `23/11/2026` | Dates |
| Texte | `E001` | Matricules et codes |

### 6.2. Appliquer un format

1. Sélectionnez la cellule ou la plage.
2. Allez dans l'onglet **Accueil**.
3. Dans le groupe **Nombre**, ouvrez la liste des formats.
4. Choisissez le format adapté.

**Exemple du pourcentage :** si une cellule contient `0,25`, le format Pourcentage l'affiche comme `25 %`.

**Attention :** le format Pourcentage modifie l'affichage de la valeur. Il ne faut pas confondre `0,25` affiché en pourcentage avec la saisie de `25` puis l'application du format Pourcentage, qui afficherait `2 500 %`.

### 6.3. Personnaliser l'apparence d'un tableau

Vous pouvez modifier :

- La police et sa taille.
- La couleur du texte.
- La couleur de remplissage.
- Les bordures.
- L'alignement.
- La largeur des colonnes.
- La hauteur des lignes.

Pour rendre le tableau des étudiants plus lisible :

1. Mettez les titres en gras.
2. Appliquez une couleur de fond aux en-têtes.
3. Centrez les notes.
4. Ajoutez des bordures.
5. Ajustez la largeur des colonnes.
6. Appliquez deux décimales aux moyennes si nécessaire.

**Conseil :** privilégiez une mise en forme sobre et cohérente. Un tableau doit d'abord être facile à lire.

## 7. Créer et personnaliser un graphique

Un graphique transforme un tableau de chiffres en représentation visuelle. Il permet de comparer des valeurs et de repérer rapidement des différences.

### 7.1. Choisir le bon graphique

| Type | Quand l'utiliser ? |
|---|---|
| Colonnes (histogramme) | Comparer des catégories. |
| Barres | Comparer des catégories avec des libellés longs. |
| Courbe | Montrer une évolution dans le temps. |
| Secteurs | Montrer la répartition d'un total. |

### 7.2. Exemple : comparer les moyennes des étudiants

Utilisez les colonnes **Étudiant** et **Moyenne** du tableau précédent.

1. Sélectionnez les noms des étudiants et leurs moyennes.
2. Cliquez sur **Insertion**.
3. Choisissez **Graphiques**, puis un graphique en colonnes.
4. Ajoutez le titre `Moyenne des étudiants`.
5. Vérifiez que les noms apparaissent sur l'axe horizontal.
6. Personnalisez la présentation si nécessaire.

### 7.3. Les éléments d'un graphique

- **Titre :** indique le sujet du graphique.
- **Axes :** permettent de lire les catégories et les valeurs.
- **Légende :** identifie les séries de données.
- **Étiquettes de données :** affichent les valeurs directement sur le graphique.

**À retenir :** un graphique doit être compréhensible sans longue explication. Son titre et ses données doivent être clairement identifiables.

## 8. Premières fonctions conditionnelles et fonctions complexes

Une fonction conditionnelle permet à Excel de prendre une décision à partir d'un test.

### 8.1. SI : prendre une décision

La fonction `SI` vérifie une condition. Elle renvoie un résultat si la condition est vraie et un autre si elle est fausse.

**Syntaxe :**

```excel
=SI(test_logique;valeur_si_vrai;valeur_si_faux)
```

**Exemple :** si la moyenne de l'étudiant se trouve en F2, afficher « Admis » lorsque la moyenne est au moins égale à 10.

```excel
=SI(F2>=10;"Admis";"Ajourné")
```

Décomposition :

- `F2>=10` : condition à vérifier.
- `"Admis"` : résultat si la condition est vraie.
- `"Ajourné"` : résultat si la condition est fausse.

### 8.2. ET : toutes les conditions doivent être vraies

La fonction `ET` renvoie VRAI uniquement si toutes les conditions sont vraies.

```excel
=ET(B2>=10;C2>=10)
```

Cette formule vérifie que les notes d'Informatique et de Mathématiques sont toutes les deux supérieures ou égales à 10.

### 8.3. OU : au moins une condition doit être vraie

La fonction `OU` renvoie VRAI si au moins une des conditions est vraie.

```excel
=OU(B2>=10;C2>=10)
```

Cette formule renvoie VRAI si au moins une des deux notes est supérieure ou égale à 10.

### 8.4. Combiner SI et ET

On peut associer plusieurs fonctions pour construire un test plus précis.

**Exemple :** afficher « Validé » si les notes d'Informatique et de Mathématiques sont toutes les deux au moins égales à 10.

```excel
=SI(ET(B2>=10;C2>=10);"Validé";"Non validé")
```

### 8.5. SOMME.SI : additionner selon un critère

La fonction `SOMME.SI` additionne les valeurs qui respectent un critère.

**Syntaxe :**

```excel
=SOMME.SI(plage_critère;critère;plage_somme)
```

Exemple de données :

| Filière | Montant |
|---|---:|
| Informatique | 150000 |
| Gestion | 120000 |
| Informatique | 180000 |
| Marketing | 100000 |
| Gestion | 130000 |

Pour additionner les montants de la filière Informatique :

```excel
=SOMME.SI(A2:A6;"Informatique";B2:B6)
```

**Résultat :** 330000.

## 9. Exercices d'application

Réalisez les exercices directement dans Excel. Essayez d'abord seul, puis vérifiez vos formules avec le cours.

### Exercice 1 -- Calculer les résultats d'une classe

Saisissez le tableau suivant :

| Étudiant | Informatique | Mathématiques | Anglais |
|---|---:|---:|---:|
| Fall | 15 | 12 | 10 |
| Diop | 8 | 14 | 11 |
| Ba | 17 | 16 | 13 |
| Sarr | 9 | 10 | 12 |
| Ndiaye | 14 | 13 | 16 |

**Travail demandé :**

1. Ajoutez une colonne Total.
2. Calculez le total des notes de chaque étudiant.
3. Ajoutez une colonne Moyenne.
4. Calculez la moyenne de chaque étudiant.
5. Trouvez la meilleure et la plus faible moyenne.
6. Affichez « Admis » ou « Ajourné » selon la moyenne.

### Exercice 2 -- Mettre en forme un tableau

À partir du tableau de l'exercice 1 :

1. Mettez les en-têtes en gras.
2. Appliquez une couleur de fond aux en-têtes.
3. Centrez les notes.
4. Ajoutez des bordures.
5. Ajustez la largeur des colonnes.
6. Affichez les moyennes avec deux décimales.

### Exercice 3 -- Créer un graphique

À partir du tableau de l'exercice 1 :

1. Sélectionnez les noms et les moyennes.
2. Créez un graphique en colonnes.
3. Donnez-lui le titre `Moyenne des étudiants`.
4. Vérifiez que les noms et les valeurs sont lisibles.
5. Modifiez la présentation du graphique.

### Exercice 4 -- Utiliser les fonctions conditionnelles

Reprenez le tableau des notes.

1. Utilisez `SI` pour afficher « Admis » si la moyenne est supérieure ou égale à 10.
2. Utilisez `ET` pour vérifier que les notes d'Informatique et de Mathématiques sont toutes les deux au moins égales à 10.
3. Utilisez `OU` pour vérifier qu'au moins une de ces deux notes est au moins égale à 10.

### Exercice 5 -- Calculer un total selon un critère

Saisissez les données suivantes :

| Filière | Montant |
|---|---:|
| Informatique | 150000 |
| Gestion | 120000 |
| Informatique | 180000 |
| Marketing | 100000 |
| Gestion | 130000 |

1. Calculez le total des montants de la filière Informatique avec `SOMME.SI`.
2. Calculez le total des montants de la filière Gestion.
3. Comparez les deux résultats.

## 10. Questions de compréhension

Répondez sans regarder le cours, puis vérifiez vos réponses.

1. Qu'est-ce qu'un classeur Excel ?
2. Comment identifie-t-on une cellule ?
3. Par quel symbole commence une formule ?
4. Quelle est la différence entre une formule et une fonction ?
5. À quoi sert la fonction `SOMME` ?
6. Quelle fonction permet de trouver la plus grande valeur ?
7. Quelle est la différence entre `NB` et `NBVAL` ?
8. Pourquoi utilise-t-on le format Pourcentage ?
9. Quel graphique choisir pour comparer les moyennes de plusieurs étudiants ?
10. À quoi sert la fonction `SI` ?
11. Quelle différence existe-t-il entre `ET` et `OU` ?
12. Dans quel cas utilise-t-on `SOMME.SI` ?

## 11. Corrigé indicatif des formules

Les formules ci-dessous supposent que les données commencent en ligne 2 et que les colonnes sont organisées comme suit :

- A : Étudiant
- B : Informatique
- C : Mathématiques
- D : Anglais
- E : Total
- F : Moyenne
- G : Décision

| Calcul | Formule à saisir en ligne 2 |
|---|---|
| Total | `=SOMME(B2:D2)` |
| Moyenne | `=MOYENNE(B2:D2)` |
| Décision | `=SI(F2>=10;"Admis";"Ajourné")` |
| Deux matières au moins égales à 10 | `=ET(B2>=10;C2>=10)` |
| Au moins une matière égale ou supérieure à 10 | `=OU(B2>=10;C2>=10)` |

Recopiez les formules vers le bas pour les autres étudiants.

**Résultats attendus pour les moyennes :**

| Étudiant | Moyenne | Décision |
|---|---:|---|
| Fall | 12,33 | Admis |
| Diop | 11,00 | Admis |
| Ba | 15,33 | Admis |
| Sarr | 10,33 | Admis |
| Ndiaye | 14,33 | Admis |

Pour la fonction `SOMME.SI` de l'exercice 5 :

- Informatique : `330000`
- Gestion : `250000`

## 12. Bilan de fin de journée

Avant de terminer, vous devez être capable de réaliser seul les manipulations suivantes :

- [ ] Créer et enregistrer un classeur.
- [ ] Saisir des données dans une feuille.
- [ ] Identifier une cellule et une plage.
- [ ] Effectuer des calculs avec les opérateurs.
- [ ] Utiliser `SOMME`, `MOYENNE`, `MAX`, `MIN`, `NB` et `NBVAL`.
- [ ] Appliquer des formats de cellule.
- [ ] Personnaliser un tableau.
- [ ] Créer et personnaliser un graphique.
- [ ] Utiliser `SI`, `ET`, `OU` et `SOMME.SI`.

### Pour aller plus loin

Prenez un tableau de votre choix (notes, dépenses, ventes ou paiements) et essayez de reproduire les manipulations vues pendant le cours.

Le meilleur moyen de progresser sur Excel est de pratiquer régulièrement.

---

**Fin du cours -- Jour 1**

**Prochaine séance : Jour 2 -- Organisation d'une base de données, mise en forme conditionnelle, filtres, tris et fonction RECHERCHEV.**
