/* Définition des rapports quotidiens des piliers de la riposte MVE – Haut-Uélé.
   Tirée des canevas PowerPoint des piliers (novembre 2026). Une seule source pour les formulaires,
   le tableau de bord, l'export Excel et les diapositives.

   Section :
     { id, titre, champs: [...] }                     champs simples
     { id, titre, grille: { lignes: [[id, lib, agg?]], colonnes: [[id, lib]] } }   tableau lignes × colonnes
     { id, titre, liste: { colonnes: [[id, lib, type, options?]] } }               lignes répétables (PoC, feedbacks…)
     { id, titre, textes: [[id, lib]] }               zones de texte
   Champ : [id, libellé, agg?, cle?]
     agg  : 'sum' (défaut : s'additionne entre jours et entre zones) ou 'stock' (état du jour : on garde le dernier jour)
     cle  : true = colonne du tableau de bord consolidé
   Calculs : { id, lib, f: function (v) -> nombre|null, pct?: true, cle?: true } ; v(id) lit une valeur (agrégée). */
(function (root) {
  'use strict';

  var ZONES = ['Aba', 'Boma Mangbetu', 'Doruma', 'Dungu', 'Faradje', 'Gombari', 'Isiro',
    'Makoro', 'Niangara', 'Pawa', 'Rungu', 'Wamba', 'Watsa'];
  var PROVINCE = 'Coordination provinciale';
  var HORS = 'Hors province';

  var VD = [['v', 'Vivants'], ['d', 'Décès']];
  var CS = [['c', 'Confirmés'], ['s', 'Suspects']];
  var r = function (a, b) { return a == null || !b ? null : a / b; };
  var s = function () { var t = 0; for (var i = 0; i < arguments.length; i++) t += arguments[i] || 0; return t; };

  var PILIERS = [
    {
      id: 'surv', nom: 'Surveillance', court: 'Surveillance', couleur: '#c8312d', niveau: 'zone',
      sections: [
        { id: 'cas', titre: 'Cas du jour', grille: { somme: true, lignes: [['conf', 'Nouveaux cas confirmés']], colonnes: VD } },
        { id: 'cas2', titre: 'Précisions sur les cas', champs: [
          ['conf_contact', 'dont contacts connus et suivis'], ['enf5', 'dont enfants de moins de 5 ans'],
          ['as_conf', 'dont agents de santé (PPL)'], ['dc_comm', 'Décès communautaires', 'sum', true]] },
        { id: 'al', titre: 'Situation des alertes', grille: { somme: true,
          lignes: [['rep', 'Report : alertes non vérifiées la veille', 'stock'], ['nouv', 'Nouvelles alertes reçues', 'sum', true],
            ['val', 'Vérifiées et validées (cas suspects)', 'sum', true], ['nval', 'Vérifiées et non validées']],
          colonnes: VD } },
        { id: 'al2', titre: 'Transfert des cas suspects', champs: [
          ['rep_susp', 'Report : suspects non transférés la veille', 'stock'],
          ['transf', 'Suspects investigués et transférés au CT/CTE', 'sum', true],
          ['ntransf', 'Suspects non investigués / non transférés', 'sum', true]] },
        { id: 'src', titre: 'Sources des alertes', grille: { somme: true,
          lignes: [['ess', 'ESS / recherche passive'], ['act', 'Recherche active'], ['comm', 'Communauté'], ['poe', 'PoE / PoC']], colonnes: VD } },
        { id: 'ra', titre: 'Recherche active', champs: [
          ['ra_dossiers', 'Dossiers / registres examinés'], ['ra_alertes', 'Alertes détectées dans les ESS'],
          ['ra_val', 'Alertes ESS validées (cas suspects)'], ['ra_encours', 'Alertes ESS en cours de validation', 'stock'],
          ['rc_det', 'Alertes communautaires détectées'], ['rc_ver', 'Alertes communautaires vérifiées'],
          ['rc_val', 'Alertes communautaires validées'], ['rc_dc', 'Décès communautaires suspects signalés et investigués']] },
        { id: 'sites', titre: 'Sites de recherche active', grille: {
          lignes: [['hgr', 'Hôpital général de référence'], ['ch', 'Centre hospitalier / médical'], ['cs', 'Centre de santé'],
            ['labo', 'Laboratoire'], ['ps', 'Poste de santé'], ['disp', 'Dispensaire'], ['trad', 'Tradipraticiens'],
            ['priere', 'Maisons de prière'], ['pharma', 'Pharmacies']],
          colonnes: [['plan', 'Planifiés'], ['vis', 'Visités']] } },
        { id: 'ct', titre: 'Suivi des contacts', champs: [
          ['ct_rep', 'Report (contacts en suivi la veille)', 'stock'], ['ct_nouv', 'Nouveaux contacts listés', 'sum', true],
          ['ct_vus', 'Contacts vus', 'sum', true], ['ct_sortis', 'Sortis (21 jours terminés)'], ['ct_susp', 'Devenus suspects', 'sum', true],
          ['ct_conf', 'Devenus confirmés'], ['nv_abs', 'Non vus : simple absence'], ['nv_pdv', 'Non vus : perdus de vue (≥ 72 h)'],
          ['nv_jamais', 'Non vus : jamais vus'], ['nv_refus', 'Non vus : refus'], ['nv_depl', 'Non vus : déplacés']] }
      ],
      calculs: [
        { id: 'conf_tot', lib: 'Nouveaux cas confirmés', f: function (v) { return s(v('cas.conf.v'), v('cas.conf.d')); }, cle: true, avant: true },
        { id: 'al_tot', lib: 'Total alertes (report + nouvelles)', f: function (v) { return s(v('al.rep.v'), v('al.rep.d'), v('al.nouv.v'), v('al.nouv.d')); }, cle: true },
        { id: 'al_valtot', lib: 'Alertes validées', f: function (v) { return s(v('al.val.v'), v('al.val.d')); } },
        { id: 'p_isol', lib: '% suspects validés transférés au CT/CTE', pct: true, f: function (v) { return r(v('transf'), s(v('al.val.v'), v('al.val.d'))); }, cle: true },
        { id: 'p_sites', lib: '% sites de recherche active visités', pct: true, f: function (v) {
          var p = 0, x = 0; ['hgr', 'ch', 'cs', 'labo', 'ps', 'disp', 'trad', 'priere', 'pharma'].forEach(function (k) { p += v('sites.' + k + '.plan') || 0; x += v('sites.' + k + '.vis') || 0; });
          return r(x, p); } },
        { id: 'ct_asuivre', lib: 'Contacts à suivre ce jour', f: function (v) { return s(v('ct_rep'), v('ct_nouv')); }, cle: true },
        { id: 'ct_taux', lib: 'Taux de suivi des contacts', pct: true, f: function (v) { return r(v('ct_vus'), s(v('ct_rep'), v('ct_nouv'))); }, cle: true },
        { id: 'ct_nonvus', lib: 'Contacts non vus', f: function (v) { var a = s(v('ct_rep'), v('ct_nouv')); return a ? a - (v('ct_vus') || 0) : null; } },
        { id: 'ct_reste', lib: 'Reste à suivre le jour suivant', f: function (v) { var a = s(v('ct_rep'), v('ct_nouv')); return a ? a - s(v('ct_sortis'), v('ct_susp'), v('ct_conf')) : null; } }
      ]
    },
    {
      id: 'pec', nom: 'Prise en charge holistique', court: 'Prise en charge', couleur: '#1d4a6b', niveau: 'zone', unite: 'CTE / CT',
      entete: [['unite', 'CTE / CT', 'texte']],
      sections: [
        { id: 'lits', titre: 'Lits montés', grille: { somme: true, lignes: [['lits', 'Lits montés', 'stock']], colonnes: CS } },
        { id: 'mv', titre: 'Mouvements des malades', grille: { somme: true,
          lignes: [['rep', 'Reports (veille)', 'stock'], ['adm', 'Nouvelles admissions', 'sum', true], ['adm_ppl', 'dont PPL admis'],
            ['hosp', 'Hospitalisés (cas actifs) en fin de journée', 'stock', true], ['fe', 'dont femmes enceintes', 'stock'],
            ['fa', 'dont femmes allaitantes', 'stock'], ['m5', 'dont enfants < 5 ans', 'stock']], colonnes: CS } },
        { id: 'prel', titre: 'Laboratoire au CTE', champs: [['prel', 'Prélèvements réalisés sur les nouveaux admis (suspects)']] },
        { id: 'so', titre: 'Sorties', grille: { somme: true,
          lignes: [['gueris', 'Guéris', 'sum', true], ['nc', 'Sortis non-cas'], ['dec', 'Décès', 'sum', true], ['dec24', 'dont décès < 24 h après admission'],
            ['evad', 'Évadés', 'sum', true], ['transf', 'Transférés']], colonnes: CS } },
        { id: 'st', titre: 'Stade clinique des hospitalisés', grille: { somme: true,
          lignes: [['leger', 'Léger', 'stock'], ['modere', 'Modéré', 'stock'], ['severe', 'Sévère / critique', 'stock']], colonnes: CS } },
        { id: 'nut', titre: 'Prise en charge nutritionnelle', grille: {
          lignes: [['regime', 'Régime spécifique'], ['mas', 'Malnutrition aiguë sévère (MAS)'], ['mam', 'Malnutrition aiguë modérée (MAM)'], ['repas', 'Repas chauds servis']],
          colonnes: [['s', 'Suspects'], ['c', 'Confirmés'], ['creche', 'Crèche'], ['acc', 'Accompagnants'], ['ppl', 'PPL']] } }
      ],
      calculs: [
        { id: 'occ_c', lib: 'Taux d’occupation des lits (confirmés)', pct: true, f: function (v) { return r(v('mv.hosp.c'), v('lits.lits.c')); }, cle: true },
        { id: 'occ_s', lib: 'Taux d’occupation des lits (suspects)', pct: true, f: function (v) { return r(v('mv.hosp.s'), v('lits.lits.s')); }, cle: true }
      ]
    },
    {
      id: 'smsps', nom: 'Santé mentale et soutien psychosocial', court: 'Psychosocial', couleur: '#6b4fa0', niveau: 'zone', unite: 'CTE / CT',
      entete: [['unite', 'CTE / CT', 'texte']],
      sections: [
        { id: 'pec', titre: 'Prise en charge psychosociale', champs: [
          ['nconf', 'Nouveaux cas confirmés pris en charge', 'sum', true], ['conf_suivi', 'Cas confirmés en cours de suivi', 'stock', true],
          ['nsusp', 'Nouveaux cas suspects pris en charge', 'sum', true], ['susp_suivi', 'Cas suspects en cours de suivi', 'stock'],
          ['res', 'Résultats de laboratoire annoncés'], ['res_pos', 'dont positifs'], ['res_neg', 'dont négatifs']] },
        { id: 'pec2', titre: 'Enfants, familles et personnel', champs: [
          ['enf_creche', 'Enfants séparés en crèche', 'stock', true], ['enf_comm', 'Enfants séparés en communauté', 'stock', true],
          ['accomp', 'Accompagnants au CTE'], ['ppl', 'PPL suivis', 'sum', true], ['menages', 'Membres des ménages et communautés'],
          ['eds', 'EDS appuyés']] },
        { id: 'enf', titre: 'Situation des enfants séparés en difficulté particulière', textes: [['enf_txt', 'Situation (RAS si rien à signaler)']] }
      ],
      calculs: []
    },
    {
      id: 'crec', nom: 'Communication des risques et engagement communautaire', court: 'CREC', couleur: '#b97700', niveau: 'zone',
      sections: [
        { id: 'ind', titre: 'Indicateurs clés', grille: {
          lignes: [['al_comm', 'Alertes communautaires notifiées', 'sum', true], ['al_dec', 'Alertes décès', 'sum', true],
            ['feed', 'Feedbacks / infodémie', 'sum', true], ['inc', 'Incidents communautaires', 'sum', true], ['ref', 'Cas référés']],
          colonnes: [['recus', 'Reçus'], ['res', 'Résolus (48 h)']] } },
        { id: 'eng', titre: 'Engagement communautaire', grille: {
          lignes: [['caus', 'Causeries éducatives'], ['masse', 'Activités de masse'], ['vad', 'Visites à domicile'],
            ['plaid', 'Plaidoyers'], ['radio', 'Émissions radio']],
          colonnes: [['nb', 'Nombre'], ['h', 'Hommes sensibilisés'], ['f', 'Femmes sensibilisées']] } },
        { id: 'cas', titre: 'Activités autour des cas (appui aux autres piliers)', textes: [
          ['ac_surv', 'Surveillance (listage, suivi, transfert au CTE…)'], ['ac_pci', 'PCI / décontamination'], ['ac_eds', 'EDS'],
          ['ac_labo', 'Laboratoire'], ['ac_csse', 'Continuité des services'], ['ac_poe', 'PoE / PoC']] },
        { id: 'fb', titre: 'Feedbacks communautaires', liste: { colonnes: [
          ['texte', 'Feedback', 'texte'], ['source', 'Source et localisation', 'texte'],
          ['type', 'Type', 'choix', ['Préoccupation', 'Perception', 'Demande', 'Rumeur', 'Plainte', 'Suggestion']],
          ['pilier', 'Pilier concerné', 'texte']] } }
      ],
      calculs: [
        { id: 'sens', lib: 'Personnes sensibilisées', cle: true, f: function (v) {
          var t = 0; ['caus', 'masse', 'vad', 'plaid', 'radio'].forEach(function (k) { t += (v('eng.' + k + '.h') || 0) + (v('eng.' + k + '.f') || 0); }); return t || null; } }
      ]
    },
    {
      id: 'pci', nom: 'PCI / WASH et EDS', court: 'PCI / WASH', couleur: '#23804f', niveau: 'zone',
      sections: [
        { id: 'g', titre: 'Situation du jour', champs: [['cas_jour', 'Cas confirmés du jour pris en charge par la PCI'], ['ppl_inf', 'PPL infectés (nouveaux)', 'sum', true]] },
        { id: 'ac', titre: 'Activités autour des cas', grille: {
          lignes: [['ring', 'Rings', 'sum', true], ['men', 'Ménages', 'sum', true], ['ess', 'ESS', 'sum', true], ['egl', 'Églises'],
            ['eco', 'Écoles'], ['mar', 'Marchés'], ['aut', 'Autres lieux publics']],
          colonnes: [['cib', 'Cibles'], ['dec', 'Décontaminés / ouverts'], ['kit', 'Kits remis']] } },
        { id: 'ess', titre: 'Autres activités dans les ESS', champs: [
          ['as', 'Aires de santé couvertes'], ['ess_vis', 'ESS visités'], ['ess_eval', 'ESS évalués (scorecard)'],
          ['brief', 'Prestataires briefés'], ['triage', 'Unités de triage fonctionnelles', 'stock'], ['ess_kits', 'ESS dotés en kits PCI']] },
        { id: 'scores', titre: 'Scores et observations', textes: [['scores', 'Scores des ESS évalués et observations (ex. : CS Gossamu 41,9 %)']] },
        { id: 'eds', titre: 'Enterrements dignes et sécurisés (EDS)', champs: [
          ['eds_al_h', 'Alertes EDS : hommes'], ['eds_al_f', 'Alertes EDS : femmes'], ['eds_rep', 'Reports en fin de journée', 'stock'],
          ['eds_real', 'EDS réalisés', 'sum', true], ['eds_swab', 'Corps swabés', 'sum', true], ['eds_nonreal', 'EDS non réalisés (résistance…)']] }
      ],
      calculs: [
        { id: 'eds_al', lib: 'Alertes EDS', f: function (v) { return s(v('eds_al_h'), v('eds_al_f')) || null; }, cle: true },
        { id: 'p_ring', lib: 'Taux d’ouverture des rings', pct: true, f: function (v) { return r(v('ac.ring.dec'), v('ac.ring.cib')); }, cle: true },
        { id: 'p_men', lib: '% ménages décontaminés', pct: true, f: function (v) { return r(v('ac.men.dec'), v('ac.men.cib')); } },
        { id: 'p_eds', lib: '% alertes EDS réalisées', pct: true, f: function (v) { return r(v('eds_real'), s(v('eds_al_h'), v('eds_al_f'))); } }
      ]
    },
    {
      id: 'poe', nom: 'Points d’entrée et points de contrôle', court: 'PoE / PoC', couleur: '#0f6f7a', niveau: 'zone',
      sections: [
        { id: 'pts', titre: 'Données par PoE / PoC', liste: { colonnes: [
          ['nom', 'PoE / PoC', 'texte'], ['fonc', 'Fonctionnel sans interruption', 'choix', ['Oui', 'Non']],
          ['pass_m', 'Passés (H)', 'n'], ['pass_f', 'Passés (F)', 'n'], ['scr_m', 'Screenés (H)', 'n'], ['scr_f', 'Screenés (F)', 'n'],
          ['ref_scr', 'Refus screening', 'n'], ['lave', 'Mains lavées', 'n'], ['ref_lave', 'Refus lavage', 'n'],
          ['sens', 'Sensibilisés', 'n'], ['al_v', 'Alertes vivantes', 'n'], ['al_c', 'Corps sans vie', 'n']] } },
        { id: 'al', titre: 'Suivi des alertes', champs: [
          ['al_inv2h', 'Alertes investiguées dans les 2 h'], ['al_val', 'Alertes validées', 'sum', true],
          ['al_transf', 'Alertes transférées au CTE / EDS réalisés'], ['res24', 'Résultats disponibles en 24 h'],
          ['ct_rech', 'Contacts recherchés aux PoE / PoC'], ['listes', 'PoE / PoC avec la liste des contacts à jour', 'stock']] },
        { id: 'refus', titre: 'Retours des voyageurs', textes: [['refus_txt', 'Causes de refus évoquées par les voyageurs']] }
      ],
      calculs: [
        { id: 'pass', lib: 'Personnes passées', cle: true, f: function (v) { return s(v('pts.pass_m'), v('pts.pass_f')) || null; } },
        { id: 'p_scr', lib: '% de voyageurs screenés', pct: true, cle: true, f: function (v) { return r(s(v('pts.scr_m'), v('pts.scr_f')), s(v('pts.pass_m'), v('pts.pass_f'))); } },
        { id: 'p_lave', lib: '% ayant lavé les mains', pct: true, f: function (v) { return r(v('pts.lave'), s(v('pts.pass_m'), v('pts.pass_f'))); } },
        { id: 'al_tot', lib: 'Alertes notifiées', cle: true, f: function (v) { return s(v('pts.al_v'), v('pts.al_c')); } },
        { id: 'p_inv2h', lib: '% alertes investiguées dans les 2 h', pct: true, f: function (v) { return r(v('al_inv2h'), s(v('pts.al_v'), v('pts.al_c'))); } },
        { id: 'p_fonc', lib: '% PoE / PoC fonctionnels sans interruption', pct: true, cle: true, f: function (v) { return r(v('pts.#fonc'), v('pts.#')); } }
      ]
    },
    {
      id: 'labo', nom: 'Laboratoire et recherche', court: 'Laboratoire', couleur: '#a23b5d', niveau: 'zone', zoneLib: 'Zone de santé d’origine des échantillons',
      entete: [['labo', 'Laboratoire', 'choix', ['Isiro', 'Wamba', 'Autre']]],
      sections: [
        { id: 'ech', titre: 'Échantillons du jour', grille: {
          lignes: [['np', 'Nouveaux prélèvements – sang', 'sum', true], ['swab', 'Écouvillons (swab)', 'sum', true],
            ['r1', '1er reprélèvement'], ['r2', '2e reprélèvement'], ['r3', '3e reprélèvement']],
          colonnes: [['rec', 'Reçus'], ['ana', 'Analysés'], ['pos', 'Positifs'], ['neg', 'Négatifs']] } },
        { id: 'pos', titre: 'Positifs parmi les nouveaux prélèvements', champs: [['pos_v', 'Positifs vivants'], ['pos_d', 'Positifs décédés']] },
        { id: 'aut', titre: 'Autres', champs: [['inval', 'Échantillons invalides'], ['encours', 'Échantillons en cours d’analyse', 'stock']] }
      ],
      calculs: [
        { id: 'rec', lib: 'Échantillons reçus', cle: true, f: function (v) { return somGrille(v, 'ech', ['np', 'swab', 'r1', 'r2', 'r3'], 'rec'); } },
        { id: 'ana', lib: 'Échantillons analysés', cle: true, f: function (v) { return somGrille(v, 'ech', ['np', 'swab', 'r1', 'r2', 'r3'], 'ana'); } },
        { id: 'pos', lib: 'Résultats positifs', cle: true, f: function (v) { return somGrille(v, 'ech', ['np', 'swab', 'r1', 'r2', 'r3'], 'pos'); } },
        { id: 'neg', lib: 'Résultats négatifs', f: function (v) { return somGrille(v, 'ech', ['np', 'swab', 'r1', 'r2', 'r3'], 'neg'); } },
        { id: 'tx_pos', lib: 'Taux de positivité', pct: true, cle: true, f: function (v) {
          return r(somGrille(v, 'ech', ['np', 'swab', 'r1', 'r2', 'r3'], 'pos'), somGrille(v, 'ech', ['np', 'swab', 'r1', 'r2', 'r3'], 'ana')); } }
      ]
    },
    {
      id: 'csse', nom: 'Continuité des services de santé essentiels', court: 'Continuité des services', couleur: '#4d6478', niveau: 'les deux',
      sections: [
        { id: 'ppl', titre: 'Prestataires (PPL)', champs: [
          ['ppl_inf', 'Nouveaux PPL infectés', 'sum', true], ['ppl_suivi', 'PPL suivis au CTE', 'stock'],
          ['ppl_vacc', 'PPL vaccinés (Ervebo)', 'sum', true], ['sites_vacc', 'Sites de vaccination des PPL ouverts', 'stock', true]] },
        { id: 'ess', titre: 'Services essentiels', champs: [
          ['sent_suivis', 'ESS sentinelles suivis'], ['sent_fonc', 'ESS sentinelles fonctionnels', 'stock', true],
          ['gratuit', 'Patients pris en charge gratuitement'], ['enf_paquet', 'Enfants confirmés ayant reçu le paquet (vaccination, nutrition…)'],
          ['doses', 'Doses de vaccin reçues'], ['ucc', 'Chambres froides (UCC) fonctionnelles', 'stock']] },
        { id: 'appro', titre: 'Approvisionnement', textes: [['appro_txt', 'Vaccins et intrants (réceptions, distributions, ruptures)']] }
      ],
      calculs: [
        { id: 'p_sent', lib: '% ESS sentinelles fonctionnels', pct: true, f: function (v) { return r(v('sent_fonc'), v('sent_suivis')); } }
      ]
    },
    {
      id: 'log', nom: 'Logistique', court: 'Logistique', couleur: '#7a5a2b', niveau: 'les deux',
      sections: [
        { id: 'carb', titre: 'Carburant (litres)', champs: [['carb_ant', 'Solde antérieur', 'stock'], ['carb_in', 'Entrées'], ['carb_out', 'Sorties', 'sum', true]] },
        { id: 'parc', titre: 'Transport et énergie', champs: [
          ['veh', 'Véhicules disponibles', 'stock', true], ['veh_piliers', 'Véhicules mis à disposition des piliers ce jour'],
          ['amb', 'Ambulances fonctionnelles', 'stock', true], ['motos', 'Motos / tricycles disponibles', 'stock'],
          ['ge', 'Groupes électrogènes fonctionnels', 'stock']] },
        { id: 'volets', titre: 'Activités par volet', textes: [
          ['stocks_txt', 'Gestion des stocks (réceptions, livraisons aux piliers)'], ['infra_txt', 'Infrastructure (CTE, UMS…)'],
          ['transp_txt', 'Transport, charroi automobile et groupes électrogènes']] }
      ],
      calculs: [
        { id: 'carb_solde', lib: 'Solde de carburant (L)', cle: true, f: function (v) {
          var a = v('carb_ant'), e = v('carb_in'), o = v('carb_out'); return a == null && e == null && o == null ? null : s(a, e) - (o || 0); } }
      ]
    }
  ];
  function somGrille(v, sec, lignes, col) {
    var t = null; lignes.forEach(function (l) { var x = v(sec + '.' + l + '.' + col); if (x != null) t = (t || 0) + x; }); return t;
  }

  // Sections communes à tous les rapports (comme dans chaque présentation)
  var COMMUN_DEBUT = [{ id: 'saillants', titre: 'Points saillants du jour', textes: [['saillants', 'Points saillants (RAS si rien à signaler)']] }];
  var COMMUN_FIN = [
    { id: 'realisees', titre: 'Activités réalisées', textes: [['realisees', 'Activités réalisées']] },
    { id: 'photos', titre: 'Photos des activités', photos: true },
    { id: 'defis', titre: 'Défis et actions correctrices', liste: { colonnes: [
      ['defi', 'Défi / problème', 'texte'], ['solution', 'Solution proposée', 'texte'], ['resp', 'Responsable', 'texte'], ['delai', 'Délai', 'texte']] } },
    { id: 'planifiees', titre: 'Activités planifiées', textes: [['planifiees', 'Activités planifiées pour demain']] }
  ];

  // PoE / PoC connus (canevas du 08/11/2026), proposés par zone
  var POINTS = {
    'Makoro': ['Boh'], 'Watsa': ['PoE de Doko'],
    'Isiro': ['Aéroport Matari', 'PK10 route Poko', 'PK6 route Niangara', 'Magambe', 'Gossamu'],
    'Wamba': ['Salisa', 'Mutowamba', 'Bombombi', 'PK18'], 'Boma Mangbetu': ['Bangwene'], 'Pawa': ['Bolebole']
  };

  /* Liste à plat des champs numériques d'un pilier : { id, lib, sec, agg, cle } */
  function champs(p) {
    var out = [];
    p.sections.forEach(function (sec) {
      (sec.champs || []).forEach(function (c) { out.push({ id: c[0], lib: c[1], sec: sec, agg: c[2] || 'sum', cle: !!c[3] }); });
      if (sec.grille) sec.grille.lignes.forEach(function (l) {
        sec.grille.colonnes.forEach(function (c) {
          out.push({ id: sec.id + '.' + l[0] + '.' + c[0], lib: l[1] + ' – ' + c[1], sec: sec, agg: l[2] || 'sum', cle: !!l[3], ligne: l, col: c });
        });
      });
      if (sec.liste) sec.liste.colonnes.forEach(function (c) {
        if (c[2] === 'n') out.push({ id: sec.id + '.' + c[0], lib: c[1] + ' (total)', sec: sec, agg: 'sum', liste: true });
      });
    });
    return out;
  }
  /* Formulaire des valeurs de départ : seulement ce qui s'additionne (pas de « fin de journée », pas de textes) */
  function cumul(t) { return t.replace(/ du jour$/, '') + ' – cumul'; }
  function sectionsDepart(p) {
    var out = [];
    p.sections.forEach(function (sec) {
      if (sec.champs) { var ch = sec.champs.filter(function (c) { return (c[2] || 'sum') !== 'stock'; }); if (ch.length) out.push({ id: sec.id, titre: cumul(sec.titre), champs: ch }); }
      if (sec.grille) { var l = sec.grille.lignes.filter(function (x) { return (x[2] || 'sum') !== 'stock'; }); if (l.length) out.push({ id: sec.id, titre: cumul(sec.titre), grille: { lignes: l, colonnes: sec.grille.colonnes, somme: sec.grille.somme } }); }
      if (sec.liste) { var n = sec.liste.colonnes.filter(function (c) { return c[2] === 'n'; }); if (n.length) out.push({ id: sec.id + '-tot', titre: cumul(sec.titre), champs: n.map(function (c) { return [sec.id + '.' + c[0], c[1]]; }) }); }
    });
    return out;
  }
  function pilier(id) { for (var i = 0; i < PILIERS.length; i++) if (PILIERS[i].id === id) return PILIERS[i]; return null; }
  function sections(p) { return COMMUN_DEBUT.concat(p.sections, COMMUN_FIN); }
  function zonesPour(p) {
    if (p.niveau === 'les deux') return [PROVINCE].concat(ZONES);
    if (p.id === 'labo') return ZONES.concat([HORS]);
    return ZONES.concat([PROVINCE]);
  }

  var api = { ZONES: ZONES, PROVINCE: PROVINCE, HORS: HORS, PILIERS: PILIERS, POINTS: POINTS, COMMUN_DEBUT: COMMUN_DEBUT, COMMUN_FIN: COMMUN_FIN,
    champs: champs, pilier: pilier, sections: sections, sectionsDepart: sectionsDepart, zonesPour: zonesPour };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.PILIERS_DEF = api;
})(this);
