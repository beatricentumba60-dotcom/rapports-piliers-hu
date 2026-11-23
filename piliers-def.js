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
    { id: 'realisees', titre: 'Activités réalisées', textes: [['realisees', 'Activités réalisées (une activité par ligne)']] },
    { id: 'photos', titre: 'Photos des activités', photos: true },
    { id: 'defis', titre: 'Défis et pistes de solution', liste: { colonnes: [
      ['defi', 'Défi / problème', 'texte'], ['solution', 'Piste de solution', 'texte'], ['resp', 'Responsable', 'texte'], ['delai', 'Délai', 'texte']] } },
    { id: 'planifiees', titre: 'Activités planifiées', textes: [['planifiees', 'Activités planifiées pour demain (une par ligne)']] }
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

  /* Aires de santé par zone et population 2026 (DHIS2 SNIS RDC, extraction du 12/11/2026 ; 213 aires) */
  var AIRES = {
    Aba: [["Aba", 37803], ["Atadra", 8140], ["Ataki", 7259], ["Baki", 6484], ["Bele Patarua", 12636], ["Dramba", 18791], ["Kamiru", 6996], ["Kitambala", 23380], ["Kurukwata", 9687], ["Lagabe", 12986], ["Lipara", 11622], ["Massabe", 10072], ["Nyalanya", 34642], ["Nyari", 5458], ["Tayo", 11929], ["Todro", 8587]],
    "Boma Mangbetu": [["Adedo", 7432], ["Adjangwe", 9679], ["Badamoni", 11360], ["Bafwabaka", 10509], ["Bakangi", 9219], ["Bavamasia", 9967], ["Bevesieni", 10508], ["Boma Mangbetu", 13531], ["Gombe", 11644], ["Legu", 10326], ["Mataku", 8664], ["Nambose", 9621], ["Obongoni", 15323], ["Yamo", 14186]],
    Doruma: [["Bangadi", 23957], ["Diagbe", 6341], ["Diebio", 1718], ["Doruma", 17684], ["Gangala", 9821], ["Kana", 1558], ["Kapili", 1287], ["Masombo", 5033], ["Naparka", 5103], ["Napopo", 2763], ["Ngilima", 17594], ["Weneki", 3774]],
    Dungu: [["Afu", 5679], ["Bakote", 20153], ["Bamokandi", 13138], ["Biodi", 4437], ["Bitima", 6979], ["Dungu May", 4812], ["Dungu Uye", 19285], ["Duru", 5037], ["Kiliwa", 4387], ["Kosa", 13675], ["Kpekpele", 7163], ["Li Pay", 7092], ["Li Uye", 7040], ["Mangese", 25314], ["Mbadainga", 5489], ["Musa", 16772], ["Nango", 5203], ["Ndedu", 8941], ["Sambia", 30933], ["Wandote", 25467]],
    Faradje: [["Akua", 9619], ["CECA 20 Faradje", 10367], ["CECA 20 Tadu", 16507], ["Djabir", 10550], ["Doya", 8735], ["Faradje", 19354], ["Katanga", 7798], ["Kirima", 4780], ["Koreri", 4612], ["Kpodo", 10497], ["Kurukwata", 10389], ["Nzopi", 9949], ["Obi", 7336], ["Sabuni", 9916], ["Sesenge", 10199], ["Tadu", 12139], ["Tandro", 11079], ["Tomati", 6309]],
    Gombari: [["Abakudu", 5191], ["Akpandau", 9346], ["Apodo", 12057], ["Baitebi", 9303], ["Bakiri", 9998], ["Gombari", 13473], ["Goria", 22725], ["Kossia", 5719], ["Madja", 9250], ["Nepoko", 5085], ["Ngazizi", 8513], ["Tangi", 12537], ["Tibodri", 8907], ["Tora", 12530], ["Wanga", 22892]],
    Isiro: [["Arindru", 10001], ["Bunie", 13172], ["Djombe", 11279], ["Egbunda", 12341], ["Gamba", 27072], ["Gossamu", 20439], ["Makpulu", 17612], ["Mayogo", 56168], ["Mbene", 11684], ["Medje", 28427], ["N'sele", 25360], ["Nagwa", 9384], ["Nangosira", 10668], ["Nduka", 9133], ["Neisu", 26316], ["Nembombi", 14122], ["Nolua", 13725], ["Telalingbi", 12950], ["Tely", 36219], ["Tuluba", 32849]],
    Makoro: [["Abimva", 27884], ["Ambarau", 14952], ["Bhamo", 23486], ["Bovi", 37383], ["Kamiro", 15540], ["Kana", 16386], ["Kilima", 12222], ["Kiraka", 15572], ["Lanza", 21596], ["Makoro", 35869], ["Mariaba", 13911], ["Maulo", 6206], ["Ndolomo", 22048], ["Ngili", 11720], ["Rangu", 7857]],
    Niangara: [["Arikazi", 8741], ["Babagu", 12304], ["Ekibondo", 10461], ["Galagala", 9928], ["Ganga", 10209], ["Ligunza", 12993], ["Lipombo", 14102], ["Makilingbo", 8491], ["Mangada", 6436], ["Mangalu", 9547], ["Nambia", 5198], ["Nangbama", 7707], ["Ndingba", 10167], ["Tapili", 9962], ["Wamba Moke", 10577], ["Wawe", 8666]],
    Pawa: [["Abiangama", 12456], ["Babia", 10881], ["Babonde", 14811], ["Bachenze", 8820], ["Bamoka", 11036], ["Bavamabutu", 10433], ["Budubudu", 9329], ["Digi", 9681], ["Etakamokongo", 10701], ["Fungula", 13908], ["Gatoa", 11121], ["Gbonzunzu", 12512], ["Ibambi Cecca 16", 16365], ["Ibambi Etat", 15122], ["Likasi", 10298], ["Matibika", 11751], ["Nebobongo", 18991], ["Pawa", 18886]],
    Rungu: [["Badolo", 9527], ["Bala", 14881], ["Bodjo", 9780], ["Kabome", 10600], ["Kpetele", 12610], ["Mambunga", 9680], ["Nala", 9465], ["Nangazizi", 16458], ["Nekalagba", 7627], ["Nepomeda", 11352], ["Nesira", 11558], ["Ngabi", 10495], ["Rungu", 14541], ["Vube", 12775]],
    Wamba: [["Abanagomu", 12399], ["Abiba", 12347], ["Agbau", 11479], ["Asamboa", 8100], ["Bakandja", 25052], ["Bakhita", 7052], ["Bavemo", 9773], ["Bayenga", 14363], ["Betongwe", 12616], ["Bindani Bagoya", 9174], ["Bole Bole", 22122], ["Bomana", 8130], ["CECCA 16", 8706], ["Maboma", 10135], ["Mama Kayenga", 12230], ["Mambati", 16420], ["Tshaka Tshaka", 9475]],
    Watsa: [["Arumbi", 29198], ["Bethanie", 43188], ["Dilolo", 24969], ["Doko", 32229], ["Dubele", 24968], ["Durba", 28438], ["Gambela", 21201], ["Ganoza", 33412], ["Gatanga", 32289], ["Kibali", 44941], ["Kokoro", 14014], ["Maba", 8616], ["Maitulu", 27807], ["Moku", 32977], ["Monama", 30610], ["Monya", 29205], ["Ngangazo", 24072], ["Toyota", 72736]]
  };
  function aires(zone) { return (AIRES[zone] || []).map(function (a) { return a[0]; }); }
  function populationAire(zone, aire) { var a = (AIRES[zone] || []).find(function (x) { return x[0] === aire; }); return a ? a[1] : null; }

  var api = { ZONES: ZONES, PROVINCE: PROVINCE, HORS: HORS, PILIERS: PILIERS, POINTS: POINTS, COMMUN_DEBUT: COMMUN_DEBUT, COMMUN_FIN: COMMUN_FIN, AIRES: AIRES, aires: aires, populationAire: populationAire,
    champs: champs, pilier: pilier, sections: sections, sectionsDepart: sectionsDepart, zonesPour: zonesPour };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.PILIERS_DEF = api;
})(this);
