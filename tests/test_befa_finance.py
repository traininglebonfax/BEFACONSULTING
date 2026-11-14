"""
Lot 1 BEFA - moteur financier.

Les jeux de reference sont tires des dossiers reels du cabinet
("Model Business Plan/") :
  - COREAL FOOD (nouveau projet) : compte d'exploitation sur 6 ans, plan de
    financement (annee 0 / annee 1), ratio de couverture ;
  - O'IVOIRE (entreprise existante) : analyse du bilan et du compte de
    resultat 2025 / 2024.
Les echeanciers de prets sont verifies contre les formules bancaires
standard (les frais financiers des dossiers ne sont pas reproductibles a
partir des conditions de pret annoncees dans ces memes dossiers).
"""
import pytest

from befa.historique import AnalyseHistorique, BilanCondense, CompteResultatCondense, PosteImmobilisation
from befa.prets import FinanceInputError, Pret
from befa.projection import (
    Apport, Investissement, LigneCharge, LigneProduit, ProjectionInput, calculer_projection,
)

ANNEES6 = [f"Année {i}" for i in range(1, 7)]


def coreal_input(**extra) -> ProjectionInput:
    regimes = [114_400, 143_000, 171_600, 171_600, 171_600, 171_600]
    kwargs = dict(
        annees=ANNEES6,
        produits=[
            LigneProduit("Huile de palme brute", quantites=[25_168, 31_460] + [37_752] * 4, prix_unitaire=620_000, unite="t"),
            LigneProduit("Amande de palmiste", quantites=[4_004, 5_005] + [6_006] * 4, prix_unitaire=120_000, unite="t"),
        ],
        charges=[
            LigneCharge("Régimes de palme (achat + déchargement + transport)", "achats_matieres", quantites=regimes, cout_unitaire=95_000),
            LigneCharge("Autres achats", "autres_achats", pct_ca=0.05),
            LigneCharge("Transport", "transport", montants=[723_808_800] + [904_761_000] * 5),
            LigneCharge("Services extérieurs", "services_exterieurs", montants=[691_639_520] + [864_549_400] * 5),
            LigneCharge("Impôts et taxes", "impots_taxes",
                        montants=[14_315_333, 17_894_167, 17_894_167, 17_894_171, 17_894_171, 17_894_175]),
            LigneCharge("Salaires", "salaires", pct_ca=0.05),
        ],
        dotations_force=[260_000_000] * 6,
        autres_frais_financiers=[207_709_377, 168_842_918, 140_964_974, 110_622_875, 79_397_807, 50_966_514],
        impots_sur_benefice_force=[481_494_625, 516_917_828, 516_917_828, 548_097_495, 548_097_495, 569_417_980],
    )
    kwargs.update(extra)
    return ProjectionInput(**kwargs)


# ----------------------------------------------------------------------
# Compte d'exploitation COREAL (page 52 du dossier)
# ----------------------------------------------------------------------
def test_coreal_compte_exploitation_reproduit_au_franc_pres():
    ce = calculer_projection(coreal_input()).compte
    assert ce.chiffre_affaires == [16_084_640_000, 20_105_800_000] + [24_126_960_000] * 4
    assert ce.charges_par_rubrique["achats_matieres"] == [10_868_000_000, 13_585_000_000] + [16_302_000_000] * 4
    assert ce.marge_brute == [5_216_640_000, 6_520_800_000] + [7_824_960_000] * 4
    assert ce.charges_par_rubrique["autres_achats"] == [804_232_000, 1_005_290_000] + [1_206_348_000] * 4
    assert ce.valeur_ajoutee == [2_982_644_347, 3_728_305_433, 4_831_407_433, 4_831_407_429, 4_831_407_429, 4_831_407_425]
    assert ce.ebe == [2_178_412_347, 2_723_015_433, 3_625_059_433, 3_625_059_429, 3_625_059_429, 3_625_059_425]
    assert ce.resultat_exploitation == [1_918_412_347, 2_463_015_433, 3_365_059_433, 3_365_059_429, 3_365_059_429, 3_365_059_425]
    assert ce.resultat_net == [1_229_208_345, 1_777_254_687, 2_707_176_631, 2_706_339_059, 2_737_564_127, 2_744_674_931]
    assert ce.caf == [1_489_208_345, 2_037_254_687, 2_967_176_631, 2_966_339_059, 2_997_564_127, 3_004_674_931]


def test_coreal_ratio_de_couverture():
    cv = calculer_projection(coreal_input()).couverture
    assert [round(r) for r in cv.ratio_couverture] == [7, 12, 21, 27, 38, 59]
    assert cv.free_cash_flow == calculer_projection(coreal_input()).compte.caf


def test_coreal_plan_de_financement_annee0_et_annee1():
    """Annee 0 : CMT + apport finance l'investissement ; annee 1 : credit-bail
    pour le materiel roulant ; le credit spot (campagne) reste hors plan."""
    entree = coreal_input(
        autres_frais_financiers=None,
        impots_sur_benefice_force=None,
        investissements=[
            Investissement("Usine (hors matériel roulant)", 2_572_002_260, annee=0),
            Investissement("Matériel de transport et manutention", 128_700_000, annee=1),
        ],
        prets=[
            Pret("CMT", 1_742_248_634, 0.09, 84, differe_total_mois=12, differe_partiel_mois=6,
                 periodicite="trimestrielle", annee_deblocage=0),
            Pret("Crédit-bail", 128_700_000, 0.09, 48, type="CREDIT_BAIL", annee_deblocage=1),
            Pret("Spot", 1_040_000_000, 0.09, 12, type="SPOT", in_fine=True, annee_deblocage=1),
        ],
        apports=[Apport("Fonds propres (Apport du promoteur)", 829_753_626, annee=0)],
    )
    proj = calculer_projection(entree)
    plan = proj.plan
    assert plan.colonnes[0] == "Année 0" and len(plan.colonnes) == 7
    assert plan.flux_entrants[0] == 2_572_002_260  # chiffre du dossier
    assert plan.lignes_emplois[0].valeurs[:2] == [2_572_002_260, 128_700_000]
    assert plan.flux_entrants[1] == proj.compte.caf[0] + 128_700_000
    libelles = [l.libelle for l in plan.lignes_ressources]
    assert not any("Spot" in l for l in libelles)
    # identites comptables
    for c in range(7):
        assert plan.flux_nets[c] == plan.flux_entrants[c] - plan.flux_sortants[c]
    assert plan.cumul[-1] == sum(plan.flux_nets)
    # le spot pese bien sur les frais financiers (93,6 M d'interets la 1re annee)
    assert proj.echeanciers["Spot"]["interets"][0] == pytest.approx(93_600_000)
    # IS calcule a 25 % du resultat avant impot
    ce = proj.compte
    assert ce.impot_sur_benefice[0] == round(ce.resultat_avant_impot[0] * 0.25)


def test_dotations_calculees_depuis_les_investissements():
    entree = ProjectionInput(
        annees=["2026", "2027", "2028"],
        produits=[LigneProduit("Prestations", montants=[100_000_000] * 3)],
        investissements=[
            Investissement("Terrain", 20_000_000, annee=0),  # non amortissable
            Investissement("Foreuse", 30_000_000, annee=1, duree_amortissement_ans=5),
            Investissement("Véhicule", 12_000_000, annee=2, duree_amortissement_ans=4),
        ],
        dotations_existantes=[1_000_000] * 3,
    )
    ce = calculer_projection(entree).compte
    assert ce.dotations == [7_000_000, 10_000_000, 10_000_000]


def test_bfr_en_jours():
    entree = ProjectionInput(
        annees=["A1", "A2"],
        produits=[LigneProduit("Ventes", montants=[360_000_000, 720_000_000])],
        charges=[LigneCharge("Matières", "achats_matieres", pct_ca=0.5)],
        jours_clients=30, jours_stocks=10, jours_fournisseurs=20,
    )
    p = calculer_projection(entree)
    # A1 : 30 M clients + 5 M stocks - 10 M fournisseurs = 25 M ; A2 : double
    assert p.bfr == [25_000_000, 50_000_000]
    assert p.augmentation_bfr == [25_000_000, 25_000_000]
    assert any(l.libelle == "Augmentation du BFR" for l in p.plan.lignes_emplois)


def test_validation_messages_clairs():
    with pytest.raises(FinanceInputError, match="une par année"):
        calculer_projection(ProjectionInput(annees=["A1", "A2"], produits=[LigneProduit("X", montants=[1, 2, 3])]))
    with pytest.raises(FinanceInputError, match="rubrique inconnue"):
        calculer_projection(ProjectionInput(annees=["A1"], produits=[LigneProduit("X", montants=[1])],
                                            charges=[LigneCharge("Y", "loyers", montants=1)]))
    with pytest.raises(FinanceInputError, match="un seul mode"):
        calculer_projection(ProjectionInput(annees=["A1"], produits=[LigneProduit("X", montants=[1])],
                                            charges=[LigneCharge("Y", "transport", montants=1, pct_ca=0.1)]))


# ----------------------------------------------------------------------
# Echeanciers de prets
# ----------------------------------------------------------------------
def test_pret_annuites_constantes_sans_differe():
    ech = Pret("CMT", 1_000_000, 0.12, 12).echeancier()
    assert len(ech) == 12
    assert ech[0].echeance == pytest.approx(88_848.79, abs=0.01)
    assert all(e.echeance == pytest.approx(ech[0].echeance) for e in ech)
    assert sum(e.amortissement for e in ech) == pytest.approx(1_000_000)
    assert ech[-1].capital_fin == pytest.approx(0, abs=1e-6)


def test_pret_differe_partiel_interets_seuls():
    ech = Pret("CMT", 122_475_094, 0.105, 36, differe_partiel_mois=4).echeancier()
    interet_mensuel = 122_475_094 * 0.105 / 12
    for e in ech[:4]:
        assert e.amortissement == 0 and e.echeance == pytest.approx(interet_mensuel)
    assert ech[4].amortissement > 0
    assert sum(e.amortissement for e in ech) == pytest.approx(122_475_094)


def test_pret_differe_total_capitalise_les_interets():
    p = Pret("CMT", 1_000_000, 0.12, 24, differe_total_mois=3, periodicite="trimestrielle")
    ech = p.echeancier()
    assert len(ech) == 8
    assert ech[0].echeance == 0 and ech[0].interet == pytest.approx(30_000)
    assert ech[1].capital_debut == pytest.approx(1_030_000)
    assert sum(e.amortissement for e in ech) == pytest.approx(1_030_000)


def test_pret_in_fine():
    ech = Pret("Spot", 1_040_000_000, 0.09, 12, type="SPOT", in_fine=True).echeancier()
    assert all(e.amortissement == 0 for e in ech[:-1])
    assert ech[-1].amortissement == pytest.approx(1_040_000_000)
    assert sum(e.interet for e in ech) == pytest.approx(93_600_000)


def test_pret_agregation_annuelle_et_hors_horizon():
    p = Pret("CMT", 1_742_248_634, 0.09, 84, differe_total_mois=12, differe_partiel_mois=6, periodicite="trimestrielle")
    agg = p.par_annee(6)
    total_amorti = sum(agg["amortissements"]) + agg["hors_horizon"]
    capital_capitalise = p.echeancier()[4].capital_debut  # apres 4 trimestres de differe total
    assert total_amorti == pytest.approx(capital_capitalise)
    assert agg["amortissements"][0] == 0  # 12 mois de differe total
    assert agg["hors_horizon"] > 0  # 84 mois > 6 ans
    assert agg["capital_fin"][-1] == pytest.approx(agg["hors_horizon"])


def test_pret_decalage_et_annee_deblocage():
    p = Pret("CMT", 1_200_000, 0.0, 12, annee_deblocage=2, decalage_mois=6)
    agg = p.par_annee(3)
    assert agg["amortissements"] == pytest.approx([0, 600_000, 600_000])
    assert agg["capital_fin"] == pytest.approx([0, 600_000, 0])


def test_pret_validation():
    errs = Pret("CMT", 1_000, 0.09, 10, periodicite="trimestrielle").validate(1)
    assert any("multiple de la périodicité" in e for e in errs)
    errs = Pret("CMT", 1_000, 0.09, 12, differe_total_mois=12).validate(1)
    assert any("différé" in e for e in errs)
    with pytest.raises(FinanceInputError):
        Pret("CMT", -5, 0.09, 12).echeancier()


# ----------------------------------------------------------------------
# Analyse historique O'IVOIRE (pages 8 et 9 du dossier)
# ----------------------------------------------------------------------
def oivoire_bilan(exercice, amort, stocks, avances, clients, treso, report, rn, fourn, fisc, soc):
    return BilanCondense(
        exercice=exercice,
        immobilisations=[
            PosteImmobilisation("Frais d'établissement", 250_000, amort[0]),
            PosteImmobilisation("Installation & Agencement", 7_589_600, amort[1]),
            PosteImmobilisation("Matériels", 22_508_000, amort[2]),
            PosteImmobilisation("Matériel de transport", 15_000_000, amort[3]),
            PosteImmobilisation("Autres Immobilisations", 500_000, 0),
        ],
        stocks=stocks, avances_fournisseurs=avances, creances_clients=clients, tresorerie_actif=treso,
        capital=10_000_000, reserves_report=report, resultat_net=rn, compte_courant_associes=7_000_000,
        fournisseurs=fourn, dettes_fiscales=fisc, dettes_sociales=soc,
    )


def oivoire_analyse() -> AnalyseHistorique:
    b25 = oivoire_bilan("2025", (125_000, 3_794_800, 11_254_000, 7_500_000), 73_800_190, 32_500_000, 20_230_500,
                        15_393_200, 19_966_545, 127_512_175, 549_770, 38_600, 30_600)
    b24 = oivoire_bilan("2024", (75_000, 2_276_880, 6_752_400, 4_500_000), 50_300_190, 15_000_000, 19_480_500,
                        10_300_600, 7_570_525, 101_500_815, 1_079_770, 126_500, 47_000)
    cr25 = CompteResultatCondense("2025", 537_190_400, autres_achats=61_356_330, transport=445_900,
                                  services_exterieurs=260_187_155, impots_taxes=5_186_460, salaires=28_803_600,
                                  dotations=9_069_520, impot_sur_benefice=44_629_260)
    cr24 = CompteResultatCondense("2024", 427_608_300, autres_achats=40_813_640, transport=422_400,
                                  services_exterieurs=206_823_770, impots_taxes=4_649_270, salaires=28_803_600,
                                  dotations=9_069_520, impot_sur_benefice=35_525_285)
    return AnalyseHistorique(b25, b24, cr25, cr24)


def test_oivoire_bilan_condense():
    a = oivoire_analyse()
    b = a.bilan_n
    assert b.actif_immobilise_net == 23_173_800
    assert a.bilan_n1.actif_immobilise_net == 32_243_320
    assert b.actif_circulant == 126_530_690
    assert b.total_actif == b.total_passif == 165_097_690
    assert b.capitaux_propres == 164_478_720
    assert b.passif_circulant == 618_970
    assert b.fonds_de_roulement == 141_304_920  # "141 millions"
    assert b.bfr == 125_911_720  # "126 millions"
    assert b.tresorerie_nette == 15_393_200
    r = b.ratios()
    assert r["solvabilite"] * 100 == pytest.approx(99.62, abs=0.01)  # 99,625 % exactement, affiche 99,62 % dans le dossier
    assert round(r["rentabilite_capitaux_propres"] * 100) == 78
    assert a.controles() == []


def test_oivoire_compte_de_resultat():
    a = oivoire_analyse()
    for cr, va, re_, rn, caf in ((a.cr_n, 210_014_555, 172_141_435, 127_512_175, 136_581_695),
                                 (a.cr_n1, 174_899_220, 137_026_100, 101_500_815, 110_570_335)):
        assert cr.valeur_ajoutee == va
        assert cr.resultat_exploitation == re_
        assert cr.resultat_net == rn
        assert cr.caf == caf
    lignes = {l.libelle: l.valeurs for l in a.lignes_compte_resultat()}
    assert lignes["Chiffre d'affaires"][2] == 109_582_100
    assert round(lignes["Chiffre d'affaires"][3] * 100) == 26
    bilan = {l.libelle: l.valeurs for l in a.lignes_bilan()}
    assert round(bilan["Emplois stables"][3] * 100) == -28  # baisse de 28 % citee dans le dossier


def test_controles_detectent_un_bilan_desequilibre():
    a = oivoire_analyse()
    a.bilan_n.stocks += 1_000_000
    alertes = a.controles()
    assert any("déséquilibré" in x for x in alertes)
