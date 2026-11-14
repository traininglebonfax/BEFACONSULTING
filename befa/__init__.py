"""
Business Plan au format BEFA Consulting.

Paquet independant du moteur Excel historique (engine/). Il produit un
Business Plan PDF sur le modele du cabinet BEFA (cf. dossier
"Model Business Plan/") en deux variantes :

  - "nouveau_projet"      : modele COREAL FOOD (bilan d'ouverture) ;
  - "entreprise_existante": modele O'IVOIRE (analyse bilan N / N-1).

Regle d'or : tous les chiffres sont calcules ici, de facon deterministe.
L'IA ne redige que du texte et ne fournit jamais un montant du modele.
"""
