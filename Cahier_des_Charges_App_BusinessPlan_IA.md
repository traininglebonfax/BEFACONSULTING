# Cahier des charges — Application IA de génération de Business Plans pour PME

**Document à fournir à Claude Code, avec en pièces jointes du même dossier projet :**
1. `Business_Plan_Boissons_Naturelles_CI_avec_Guide.xlsx` — le classeur modèle de référence (11 feuilles, formules vérifiées et corrigées, guide pédagogique et lexique inclus)
2. `Questionnaire_Business_Plan_PME.md` — le détail complet du questionnaire et le mapping question → cellule/feuille
3. `maquette_formulaire.html` — la maquette visuelle du formulaire allégé attendu côté interface
4. Le présent document

---

## 1. Objectif du projet

Construire une application qui permet à un porteur de PME, dans **n'importe quel secteur d'activité**, de générer automatiquement un Business Plan complet et chiffré, au format Excel, à partir d'un formulaire court. L'utilisateur ne remplit que les informations réellement spécifiques à son projet ; tout le reste (durée de projection, hypothèses macro, mise en forme, structure du modèle) est déjà figé dans un classeur modèle et ne doit jamais être redemandé.

Le parcours utilisateur attendu, dans l'ordre, est le suivant : l'utilisateur remplit le formulaire (structure définie dans `maquette_formulaire.html`) ; il clique sur « Générer mon Business Plan » ; l'application affiche un aperçu de chacune des feuilles du Business Plan directement dans l'interface, sans que rien n'ait encore été téléchargé ; l'utilisateur peut ensuite télécharger le classeur Excel complet, entièrement fonctionnel (formules vivantes, pas de valeurs figées).

Le principe directeur, qui prime sur tout le reste dans ce cahier des charges, est le suivant : **l'application ne réécrit jamais la logique financière du modèle**. Elle clone le classeur modèle fourni et n'écrit que dans les cellules d'entrée (hypothèses et données du projet) ; toutes les formules, toutes les références inter-feuilles, toute la mise en forme du modèle de référence restent strictement identiques d'un Business Plan généré à l'autre. Seuls les chiffres et libellés propres au projet changent.

## 2. Ce que l'IA doit faire, et ce qu'elle ne doit jamais faire

L'IA (Claude, appelé en arrière-plan par l'application) intervient à deux niveaux précis, et seulement à ces deux niveaux.

Premier rôle : rédiger le texte narratif que le formulaire ne demande plus à l'utilisateur — la description de l'activité et la justification du profil de croissance choisi (cf. maquette : ces deux champs portent le badge « Généré par l'IA »). Ce texte est généré à partir des réponses du formulaire (secteur, produits, prix, volumes, profil de croissance sélectionné) et inséré tel quel dans les cellules de commentaire/texte prévues à cet effet dans le classeur (notamment dans `1_HYPOTHESES`, si une zone de texte libre y est ajoutée, ou dans une note d'accompagnement séparée).

Second rôle, optionnel et à activer avec prudence : proposer des valeurs par défaut raisonnables et cohérentes avec le secteur déclaré, pour les champs que l'utilisateur laisse vides en mode simplifié (par exemple, suggérer un taux de marge courant pour un secteur donné). Si ce rôle est implémenté, chaque valeur suggérée par l'IA doit rester modifiable et doit être visuellement identifiée comme telle dans l'aperçu (badge distinct de « pré-rempli fixe » et de « saisi par l'utilisateur »).

Ce que l'IA ne doit jamais faire : calculer un chiffre d'affaires, une marge, un résultat, un flux de trésorerie, un BFR, ou toute autre grandeur financière du modèle. Tous ces calculs restent des formules Excel dans le classeur, exactement comme dans le modèle de référence. L'IA ne doit jamais non plus modifier une formule, insérer une valeur numérique codée en dur dans une cellule qui contient normalement une formule, ni renommer, déplacer ou supprimer une feuille, une ligne ou une colonne du modèle.

## 3. Contraintes techniques impératives sur le classeur Excel

Ces contraintes reprennent et généralisent les règles déjà appliquées lors de la correction du classeur modèle ; elles doivent être respectées par le moteur de génération pour chaque Business Plan produit.

Le moteur part toujours d'une copie du classeur modèle fourni (`Business_Plan_Boissons_Naturelles_CI_avec_Guide.xlsx`), jamais d'un classeur reconstruit ligne par ligne. Cela garantit que toutes les formules, mises en forme, couleurs de convention (bleu = saisie, noir = calcul, vert = lien inter-feuilles, vert clair = résultat, jaune = alerte) et liens hypertexte internes du guide et du lexique sont hérités sans risque de régression.

Seules les cellules identifiées comme cellules de saisie (remplissage bleu dans le modèle, typiquement les lignes de `1_HYPOTHESES`, les lignes produits de `2_PREVISION_CA`, les lignes de `4_INVESTISSEMENTS`) sont écrites par le moteur. Toute cellule contenant une formule dans le modèle reste une formule à l'identique dans chaque Business Plan généré — le moteur ne doit jamais faire de résolution de formule côté application puis écrire le résultat numérique à la place.

Le nombre de lignes de produits/services, de charges, ou d'investissements pouvant varier d'un projet à l'autre (formulaire à listes dynamiques), le moteur doit savoir insérer ou dupliquer proprement des lignes de saisie dans les feuilles concernées **en dupliquant aussi les formules des colonnes calculées sur ces nouvelles lignes**, et en vérifiant que les formules de totaux (sommes, `SUM`, plages nommées) qui dépendent de ce nombre de lignes sont mises à jour en conséquence. C'est le point technique le plus délicat du projet ; il correspond exactement à la classe de bug diagnostiquée et corrigée dans le modèle initial (référence codée en dur qui ne suit pas un décalage de lignes) — le moteur doit être conçu spécifiquement pour ne jamais reproduire ce défaut.

Après génération, le classeur doit être recalculé (par exemple via LibreOffice en mode headless, comme cela a été fait pour valider le modèle de référence) avant d'être proposé à l'aperçu et au téléchargement, afin que les valeurs affichées dans l'aperçu correspondent à des formules réellement évaluées et sans erreur (`#REF!`, `#DIV/0!`, `#VALUE!`). Le moteur doit vérifier ce point automatiquement (total d'erreurs de formule = 0) avant d'autoriser le téléchargement, et bloquer avec un message clair si une erreur de calcul subsiste.

Les feuilles pédagogiques du modèle (`0_GUIDE_UTILISATION` et `10_LEXIQUE_FINANCIER`) sont génériques et ne dépendent d'aucune donnée de projet : elles sont copiées à l'identique dans chaque Business Plan généré, sans aucune modification.

## 4. Le formulaire d'entrée

Le formulaire à implémenter est celui déjà maquetté dans `maquette_formulaire.html`. L'application doit reproduire fidèlement sa logique de classement en trois catégories, avec les mêmes badges visuels : champs à compléter par l'utilisateur (spécifiques au projet — secteur, produits/services avec prix/coût/volume, profil de croissance, charges fixes, investissement de départ, financement, délais clients/fournisseurs) ; champs pré-remplis par défaut et modifiables uniquement en « mode avancé » (nom du projet, forme juridique, pays/devise, durée de projection, date de démarrage, jours ouvrés, taux d'intérêt et durée d'emprunt, différé, taux d'IS, pourcentages des scénarios) ; champs générés par l'IA et absents du formulaire (description de l'activité, justification de la croissance).

Le détail champ par champ, avec pour chacun le type de contrôle (liste déroulante, bouton à bascule, chip de sélection unique, champ numérique, liste dynamique à ajout/suppression de ligne), la valeur par défaut le cas échéant, et la feuille/cellule du modèle qu'il alimente, est celui déjà documenté dans `Questionnaire_Business_Plan_PME.md`. Ce document fait foi pour le mapping ; il doit être conservé à jour si le formulaire évolue.

La validation du formulaire doit empêcher la génération tant que les champs obligatoires du mode simple ne sont pas remplis (au moins un produit/service avec prix, coût et volume ; au moins une charge fixe ; le secteur d'activité). Les champs du mode avancé, eux, ont toujours une valeur par défaut valide et ne bloquent jamais la génération s'ils ne sont pas touchés.

## 5. Aperçu avant téléchargement

Chaque feuille du classeur généré doit pouvoir être visualisée dans l'interface avant que l'utilisateur ne télécharge le fichier. Il est recommandé de convertir chaque feuille en tableau HTML pour l'aperçu (à partir du classeur recalculé, en lisant les valeurs — pas les formules — pour l'affichage), en conservant autant que possible la mise en forme du modèle (couleurs, gras des totaux, format des nombres en FCFA). L'utilisateur doit pouvoir naviguer entre les feuilles (par exemple via des onglets reproduisant l'ordre des 11 feuilles) sans recharger toute la page. Le bouton de téléchargement du fichier `.xlsx` final n'apparaît qu'une fois l'aperçu généré avec succès et sans erreur de calcul.

## 6. Architecture technique proposée

Une architecture simple en trois blocs est suffisante pour ce projet, et doit être privilégiée à une architecture plus complexe :

Un frontend (formulaire + aperçu), qui peut être une application web légère reprenant l'HTML/CSS/JS de la maquette comme base, enrichie des appels réseau nécessaires.

Un backend applicatif qui reçoit les réponses du formulaire, appelle le moteur de génération Excel, orchestre l'appel à l'IA pour les textes narratifs, déclenche le recalcul du classeur, convertit les feuilles en HTML pour l'aperçu, et sert le fichier final au téléchargement.

Un moteur de génération Excel, en Python avec `openpyxl` (cohérent avec l'outillage déjà utilisé pour corriger et enrichir le modèle de référence), responsable exclusivement de la manipulation du classeur : ouverture du modèle, écriture des cellules de saisie, gestion des lignes dynamiques, sauvegarde, puis recalcul via LibreOffice headless (même mécanisme que celui employé pour valider le fichier modèle).

## 7. Livrables attendus du développement

Un dépôt de projet structuré, contenant au minimum : le code du frontend, le code du backend, le moteur de génération Excel isolé dans son propre module (testable indépendamment de l'interface), une copie versionnée du classeur modèle de référence utilisé comme gabarit, une suite de tests qui génère un Business Plan complet à partir d'un jeu de données factice et vérifie que le fichier obtenu ne contient aucune erreur de formule et que le nombre de feuilles et leur ordre correspondent exactement au modèle de référence.

Une documentation courte expliquant comment ajouter un nouveau champ au formulaire (où le déclarer, comment le relier à une cellule du modèle) et comment mettre à jour le classeur modèle sans casser le moteur de génération.

## 8. Plan de réalisation suggéré, par lots

Premier lot : moteur de génération seul, sans interface — un script qui prend un fichier de données de projet (JSON) et produit un classeur Excel valide et recalculé, testé sur au moins trois profils de projet différents (secteur avec un seul produit, secteur avec plusieurs produits et saisonnalité, secteur avec investissement de départ lourd).

Deuxième lot : formulaire web fidèle à la maquette, connecté au moteur du premier lot, sans encore l'aperçu par feuille — génération directe du fichier téléchargeable pour valider la chaîne complète.

Troisième lot : ajout de l'aperçu par feuille dans l'interface, avant téléchargement.

Quatrième lot : intégration de l'IA pour la génération des textes narratifs (description de l'activité, justification de la croissance), et, si retenu, des suggestions de valeurs par défaut sectorielles.

## 9. Critères d'acceptation

Le Business Plan généré pour un projet de test doit s'ouvrir sans erreur dans Excel et dans LibreOffice, présenter zéro erreur de formule, conserver l'intégralité des 11 feuilles dans le même ordre que le modèle de référence, et afficher des résultats financièrement cohérents (chiffre d'affaires, EBE, résultat net et trésorerie qui évoluent de façon plausible sur la durée de projection, sans valeur aberrante liée à une ligne dupliquée ou une formule mal étendue). Le formulaire ne doit jamais demander une information déjà couverte par une valeur par défaut du mode avancé. L'aperçu affiché dans l'interface doit correspondre exactement, feuille par feuille, au contenu du fichier Excel téléchargé.
