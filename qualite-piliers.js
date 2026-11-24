/* Qualité des rapports des piliers :
   1. cohérence interne de chaque rapport ;
   2. continuité d'un jour à l'autre (le report du jour = le reste de la veille) ;
   3. concordance entre piliers pour une même zone (Surveillance ↔ Laboratoire ↔ Prise en charge ↔ PCI…).
   Gravité : 'erreur' = impossible (à corriger) ; 'alerte' = écart à vérifier / rapprocher. */
(function (root) {
  'use strict';
  var D = root.PILIERS_DEF || (typeof require !== 'undefined' ? require('./piliers-def.js') : null);
  var A = root.AGREGATS || (typeof require !== 'undefined' ? require('./agregats.js') : null);

  // Somme des valeurs présentes ; null si aucune n'est renseignée
  function S(v, ids) { var t = null; ids.forEach(function (id) { var x = v(id); if (x != null) t = (t || 0) + x; }); return t; }
  var nb = function (x) { return x == null ? '–' : (Math.round(x * 10) / 10).toLocaleString('fr-FR'); };
  // Règle a ≤ b ou a = b, évaluée seulement si les deux côtés sont renseignés
  function le(a, b) { return a == null || b == null ? null : a <= b + 1e-9; }
  function eg(a, b) { return a == null || b == null ? null : Math.abs(a - b) < 1e-9; }

  /* ---------- 1. Cohérence interne (par rapport) ---------- */
  var VD = ['v', 'd'], CS = ['c', 's'];
  var g = function (sec, l, cols) { return cols.map(function (c) { return sec + '.' + l + '.' + c; }); };
  var INTERNES = {
    surv: [
      { id: 'al', grav: 'erreur', lib: 'Alertes validées + non validées ≤ total des alertes (report + nouvelles)',
        f: function (v) { var a = S(v, g('al', 'val', VD).concat(g('al', 'nval', VD))), b = S(v, g('al', 'rep', VD).concat(g('al', 'nouv', VD))); return [le(a, b), a, b]; } },
      { id: 'transf', grav: 'alerte', lib: 'Suspects transférés + non transférés ≤ alertes validées + report des suspects',
        f: function (v) { var a = S(v, ['transf', 'ntransf']), b = S(v, g('al', 'val', VD).concat(['rep_susp'])); return [le(a, b), a, b]; } },
      { id: 'srcs', grav: 'alerte', lib: 'Total des sources des alertes = nouvelles alertes reçues',
        f: function (v) { var a = S(v, g('src', 'ess', VD).concat(g('src', 'act', VD), g('src', 'comm', VD), g('src', 'poe', VD))), b = S(v, g('al', 'nouv', VD)); return [eg(a, b), a, b]; } },
      { id: 'prec', grav: 'erreur', lib: 'Contacts connus, enfants < 5 ans et agents de santé ≤ nouveaux cas confirmés',
        f: function (v) { var b = S(v, g('cas', 'conf', VD)), a = Math.max(v('conf_contact') || 0, v('enf5') || 0, v('as_conf') || 0); return [b == null && !a ? null : le(a, b || 0), a, b]; } },
      { id: 'ra', grav: 'erreur', lib: 'Recherche active : validées ≤ vérifiées ≤ détectées',
        f: function (v) { var ok = [le(v('ra_val'), v('ra_alertes')), le(v('rc_ver'), v('rc_det')), le(v('rc_val'), v('rc_ver'))].filter(function (x) { return x != null; });
          return [ok.length ? ok.every(Boolean) : null, null, null]; } },
      { id: 'sites', grav: 'alerte', lib: 'Sites de recherche active visités ≤ planifiés',
        f: function (v) { var ko = [], n = 0; ['hgr', 'ch', 'cs', 'labo', 'ps', 'disp', 'trad', 'priere', 'pharma'].forEach(function (k) { var r = le(v('sites.' + k + '.vis'), v('sites.' + k + '.plan')); if (r != null) { n++; if (!r) ko.push(k); } });
          return [n ? !ko.length : null, null, null]; } },
      { id: 'ctvus', grav: 'erreur', lib: 'Contacts vus ≤ contacts à suivre (report + nouveaux)',
        f: function (v) { var a = v('ct_vus'), b = S(v, ['ct_rep', 'ct_nouv']); return [le(a, b), a, b]; } },
      { id: 'ctsort', grav: 'erreur', lib: 'Sortis + devenus suspects + devenus confirmés ≤ contacts à suivre',
        f: function (v) { var a = S(v, ['ct_sortis', 'ct_susp', 'ct_conf']), b = S(v, ['ct_rep', 'ct_nouv']); return [le(a, b), a, b]; } },
      { id: 'ctnv', grav: 'alerte', lib: 'Raisons de non-vus = contacts non vus (à suivre − vus)',
        f: function (v) { var a = S(v, ['nv_abs', 'nv_pdv', 'nv_jamais', 'nv_refus', 'nv_depl']), s = S(v, ['ct_rep', 'ct_nouv']);
          var b = s == null || v('ct_vus') == null ? null : s - v('ct_vus'); return [eg(a, b), a, b]; } },
      { id: 'ctconf', grav: 'alerte', lib: 'Contacts devenus confirmés ≤ nouveaux cas confirmés',
        f: function (v) { var a = v('ct_conf'), b = S(v, g('cas', 'conf', VD)); return [a ? le(a, b || 0) : null, a, b]; } }
    ],
    pec: [
      { id: 'lits', grav: 'alerte', lib: 'Hospitalisés ≤ lits montés (confirmés et suspects)',
        f: function (v) { var r = CS.map(function (c) { return le(v('mv.hosp.' + c), v('lits.lits.' + c)); }).filter(function (x) { return x != null; }); return [r.length ? r.every(Boolean) : null, S(v, g('mv', 'hosp', CS)), S(v, g('lits', 'lits', CS))]; } },
      { id: 'bilan', grav: 'alerte', lib: 'Bilan des mouvements : reports + admissions − sorties = hospitalisés (confirmés + suspects)',
        f: function (v) { var rep = S(v, g('mv', 'rep', CS)), hosp = S(v, g('mv', 'hosp', CS)); if (rep == null || hosp == null) return [null];
          var b = rep + (S(v, g('mv', 'adm', CS)) || 0) - (S(v, ['gueris', 'nc', 'dec', 'evad', 'transf'].reduce(function (t, l) { return t.concat(g('so', l, CS)); }, [])) || 0);
          return [eg(hosp, b), hosp, b]; } },
      { id: 'dont', grav: 'erreur', lib: 'Femmes enceintes, allaitantes et < 5 ans ≤ hospitalisés ; PPL admis ≤ admissions ; décès < 24 h ≤ décès',
        f: function (v) { var r = [];
          CS.forEach(function (c) { ['fe', 'fa', 'm5'].forEach(function (l) { r.push(le(v('mv.' + l + '.' + c), v('mv.hosp.' + c) || 0)); }); r.push(le(v('mv.adm_ppl.' + c), v('mv.adm.' + c) || 0)); r.push(le(v('so.dec24.' + c), v('so.dec.' + c) || 0)); });
          r = r.filter(function (x) { return x != null; }); return [r.length ? r.every(Boolean) : null, null, null]; } },
      { id: 'stades', grav: 'alerte', lib: 'Total des stades cliniques = hospitalisés',
        f: function (v) { var r = CS.map(function (c) { return eg(S(v, ['leger', 'modere', 'severe'].map(function (l) { return 'st.' + l + '.' + c; })), v('mv.hosp.' + c)); }).filter(function (x) { return x != null; });
          return [r.length ? r.every(Boolean) : null, S(v, ['leger', 'modere', 'severe'].reduce(function (t, l) { return t.concat(g('st', l, CS)); }, [])), S(v, g('mv', 'hosp', CS))]; } },
      { id: 'prel', grav: 'alerte', lib: 'Prélèvements des nouveaux admis ≤ admissions de suspects',
        f: function (v) { var a = v('prel'), b = v('mv.adm.s'); return [a ? le(a, b || 0) : null, a, b]; } },
      { id: 'sortie', grav: 'erreur', lib: 'Un confirmé ne sort pas « non-cas » ; un suspect ne sort pas « guéri »',
        f: function (v) { var a = (v('so.nc.c') || 0) + (v('so.gueris.s') || 0); return [v('so.nc.c') == null && v('so.gueris.s') == null ? null : a === 0, a, 0]; } }
    ],
    smsps: [
      { id: 'res', grav: 'alerte', lib: 'Résultats positifs + négatifs = résultats annoncés',
        f: function (v) { var a = S(v, ['res_pos', 'res_neg']), b = v('res'); return [eg(a, b), a, b]; } },
      { id: 'nconf', grav: 'alerte', lib: 'Nouveaux confirmés pris en charge ≤ confirmés en cours de suivi',
        f: function (v) { var a = v('nconf'), b = v('conf_suivi'); return [a ? le(a, b || 0) : null, a, b]; } }
    ],
    crec: [
      { id: 'res', grav: 'erreur', lib: 'Résolus ≤ reçus (alertes, feedbacks, incidents, cas référés)',
        f: function (v) { var r = ['al_comm', 'al_dec', 'feed', 'inc', 'ref'].map(function (l) { return le(v('ind.' + l + '.res'), v('ind.' + l + '.recus') || 0); }).filter(function (x) { return x != null; });
          return [r.length ? r.every(Boolean) : null, null, null]; } }
    ],
    pci: [
      { id: 'dec', grav: 'alerte', lib: 'Décontaminés / ouverts ≤ cibles (rings, ménages, ESS, lieux publics)',
        f: function (v) { var r = ['ring', 'men', 'ess', 'egl', 'eco', 'mar', 'aut'].map(function (l) { return le(v('ac.' + l + '.dec'), v('ac.' + l + '.cib')); }).filter(function (x) { return x != null; });
          return [r.length ? r.every(Boolean) : null, null, null]; } },
      { id: 'ess', grav: 'alerte', lib: 'ESS évalués ≤ ESS visités', f: function (v) { var a = v('ess_eval'), b = v('ess_vis'); return [le(a, b), a, b]; } },
      { id: 'eds', grav: 'alerte', lib: 'EDS réalisés + non réalisés ≤ alertes EDS + reports',
        f: function (v) { var a = S(v, ['eds_real', 'eds_nonreal']), b = S(v, ['eds_al_h', 'eds_al_f', 'eds_rep']); return [a ? le(a, b || 0) : null, a, b]; } },
      { id: 'swab', grav: 'alerte', lib: 'Corps swabés ≤ EDS réalisés', f: function (v) { var a = v('eds_swab'), b = v('eds_real'); return [a ? le(a, b || 0) : null, a, b]; } }
    ],
    poe: [
      { id: 'scr', grav: 'erreur', lib: 'Personnes screenées ≤ personnes passées (hommes et femmes)',
        f: function (v) { var r = [le(v('pts.scr_m'), v('pts.pass_m')), le(v('pts.scr_f'), v('pts.pass_f'))].filter(function (x) { return x != null; });
          return [r.length ? r.every(Boolean) : null, S(v, ['pts.scr_m', 'pts.scr_f']), S(v, ['pts.pass_m', 'pts.pass_f'])]; } },
      { id: 'refus', grav: 'alerte', lib: 'Refus du screening = passés − screenés',
        f: function (v) { var p = S(v, ['pts.pass_m', 'pts.pass_f']), s = S(v, ['pts.scr_m', 'pts.scr_f']), a = v('pts.ref_scr'); return [a == null || p == null || s == null ? null : eg(a, p - s), a, p == null || s == null ? null : p - s]; } },
      { id: 'lave', grav: 'alerte', lib: 'Mains lavées + refus de lavage ≤ personnes passées',
        f: function (v) { var a = S(v, ['pts.lave', 'pts.ref_lave']), b = S(v, ['pts.pass_m', 'pts.pass_f']); return [le(a, b), a, b]; } },
      { id: 'al', grav: 'erreur', lib: 'Alertes validées et investiguées dans les 2 h ≤ alertes notifiées',
        f: function (v) { var t = S(v, ['pts.al_v', 'pts.al_c']); var r = [le(v('al_val'), t || 0), le(v('al_inv2h'), t || 0)].filter(function (x) { return x != null; }); return [r.length ? r.every(Boolean) : null, Math.max(v('al_val') || 0, v('al_inv2h') || 0), t]; } }
    ],
    labo: [
      { id: 'ana', grav: 'alerte', lib: 'Échantillons analysés ≤ reçus (par type)',
        f: function (v) { var r = ['np', 'swab', 'r1', 'r2', 'r3'].map(function (l) { return le(v('ech.' + l + '.ana'), v('ech.' + l + '.rec')); }).filter(function (x) { return x != null; });
          return [r.length ? r.every(Boolean) : null, null, null]; } },
      { id: 'res', grav: 'erreur', lib: 'Positifs + négatifs ≤ analysés (par type)',
        f: function (v) { var r = ['np', 'swab', 'r1', 'r2', 'r3'].map(function (l) { var p = S(v, ['ech.' + l + '.pos', 'ech.' + l + '.neg']); return p ? le(p, v('ech.' + l + '.ana') || 0) : null; }).filter(function (x) { return x != null; });
          return [r.length ? r.every(Boolean) : null, null, null]; } },
      { id: 'posvd', grav: 'alerte', lib: 'Positifs vivants + décédés = positifs des nouveaux prélèvements de sang',
        f: function (v) { var a = S(v, ['pos_v', 'pos_d']), b = v('ech.np.pos'); return [a == null ? null : eg(a, b || 0), a, b]; } }
    ],
    csse: [
      { id: 'sent', grav: 'erreur', lib: 'ESS sentinelles fonctionnels ≤ ESS sentinelles suivis', f: function (v) { var a = v('sent_fonc'), b = v('sent_suivis'); return [le(a, b), a, b]; } }
    ],
    secu: [
      { id: 'inc', grav: 'erreur', lib: 'Agressions + véhicules + sites + menaces ≤ incidents sécuritaires',
        f: function (v) { var a = S(v, ['inc_pers', 'inc_veh', 'inc_site', 'inc_menace']), b = v('inc_tot'); return [a ? le(a, b || 0) : null, a, b]; } },
      { id: 'esc', grav: 'erreur', lib: 'Escortes réalisées ≤ escortes demandées', f: function (v) { var a = v('esc_real'), b = v('esc_dem'); return [le(a, b), a, b]; } },
      { id: 'as', grav: 'alerte', lib: 'Aires de santé inaccessibles ≤ aires de la zone',
        f: function (v, r) { var a = v('as_inacc'), n = r && D.aires(r.zone).length; return [a == null || !n ? null : le(a, n), a, n || null]; } }
    ],
    log: [
      { id: 'carb', grav: 'erreur', lib: 'Sorties de carburant ≤ solde antérieur + entrées', f: function (v) { var a = v('carb_out'), b = S(v, ['carb_ant', 'carb_in']); return [a ? le(a, b || 0) : null, a, b]; } }
    ]
  };

  /* ---------- 2. Continuité : report du jour = situation de la veille (même pilier, zone et unité) ---------- */
  var CONTINUITE = {
    surv: [{ id: 'ct', lib: 'Report des contacts = « reste à suivre » de la veille', champ: 'ct_rep', veille: function (v, K) { return K.ct_reste; } }],
    pec: [
      { id: 'repc', lib: 'Report des confirmés = hospitalisés confirmés de la veille', champ: 'mv.rep.c', veille: function (v) { return v('mv.hosp.c'); } },
      { id: 'reps', lib: 'Report des suspects = hospitalisés suspects de la veille', champ: 'mv.rep.s', veille: function (v) { return v('mv.hosp.s'); } }],
    log: [{ id: 'carb', lib: 'Solde antérieur de carburant = solde de la veille', champ: 'carb_ant', veille: function (v, K) { return K.carb_solde; } }]
  };

  /* ---------- 3. Concordance entre piliers (même zone, même période) ---------- */
  function P(pilier, lib, ids) { return { pilier: pilier, lib: lib, ids: ids }; }
  var CONCORDANCES = [
    { id: 'conf_labo', type: 'eg', lib: 'Nouveaux cas confirmés = positifs des nouveaux prélèvements (ZS d’origine)',
      a: P('surv', 'Surveillance : nouveaux confirmés', g('cas', 'conf', VD)), b: P('labo', 'Laboratoire : positifs (sang + swab)', ['ech.np.pos', 'ech.swab.pos']),
      aide: 'Un résultat rendu tard dans la nuit peut être compté le lendemain par la surveillance : vérifiez sur 7 jours.' },
    { id: 'transf_adm', type: 'eg', lib: 'Suspects transférés au CT/CTE = admissions de suspects au CT/CTE',
      a: P('surv', 'Surveillance : suspects transférés', ['transf']), b: P('pec', 'Prise en charge : suspects admis', ['mv.adm.s']),
      aide: 'Un suspect transféré vers le CTE d’une autre zone apparaît dans la zone du CTE.' },
    { id: 'prel_labo', type: 'le', lib: 'Prélèvements des nouveaux admis ≤ échantillons de sang reçus au laboratoire',
      a: P('pec', 'Prise en charge : prélèvements des admis', ['prel']), b: P('labo', 'Laboratoire : sang reçu', ['ech.np.rec']) },
    { id: 'swab', type: 'eg', lib: 'Corps swabés lors des EDS = écouvillons reçus au laboratoire',
      a: P('pci', 'PCI : corps swabés', ['eds_swab']), b: P('labo', 'Laboratoire : swabs reçus', ['ech.swab.rec']) },
    { id: 'dc_eds', type: 'le', lib: 'Décès validés (surveillance) et décès au CTE ≤ alertes EDS',
      a: { pilier: ['surv', 'pec'], lib: 'Décès : alertes validées + décès au CTE', parts: [['surv', g('al', 'val', ['d'])], ['pec', g('so', 'dec', CS)]] },
      b: P('pci', 'PCI : alertes EDS', ['eds_al_h', 'eds_al_f']), aide: 'Chaque décès suspect ou confirmé doit déclencher un EDS.' },
    { id: 'crec_dc', type: 'le', lib: 'Alertes décès reçues par la CREC ≤ nouvelles alertes décès de la surveillance',
      a: P('crec', 'CREC : alertes décès', ['ind.al_dec.recus']), b: P('surv', 'Surveillance : nouvelles alertes décès', ['al.nouv.d']),
      aide: 'Une alerte remontée par la communauté doit être transmise à la surveillance.' },
    { id: 'crec_comm', type: 'le', lib: 'Alertes communautaires (CREC) ≤ alertes de source communautaire (surveillance)',
      a: P('crec', 'CREC : alertes communautaires', ['ind.al_comm.recus']), b: P('surv', 'Surveillance : alertes de la communauté', g('src', 'comm', VD)) },
    { id: 'poe_al', type: 'eg', lib: 'Alertes des PoE/PoC = alertes de source PoE/PoC (surveillance)',
      a: P('poe', 'PoE/PoC : alertes notifiées', ['pts.al_v', 'pts.al_c']), b: P('surv', 'Surveillance : alertes des PoE/PoC', g('src', 'poe', VD)) },
    { id: 'ppl', type: 'eg', lib: 'PPL (agents de santé) infectés : surveillance = PCI',
      a: P('surv', 'Surveillance : agents de santé confirmés', ['as_conf']), b: P('pci', 'PCI : PPL infectés', ['ppl_inf']) },
    { id: 'ppl_csse', type: 'eg', lib: 'PPL infectés : PCI = continuité des services', province: true,
      a: P('pci', 'PCI : PPL infectés', ['ppl_inf']), b: P('csse', 'Continuité : PPL infectés', ['ppl_inf']) },
    { id: 'smsps_hosp', type: 'eg', lib: 'Confirmés suivis par le psychosocial = confirmés hospitalisés',
      a: P('smsps', 'Psychosocial : confirmés en suivi', ['conf_suivi']), b: P('pec', 'Prise en charge : confirmés hospitalisés', ['mv.hosp.c']),
      aide: 'Chaque confirmé hospitalisé devrait bénéficier d’un suivi psychosocial.' },
    { id: 'smsps_conf', type: 'le', lib: 'Nouveaux confirmés pris en charge (psychosocial) ≤ nouveaux confirmés (surveillance)',
      a: P('smsps', 'Psychosocial : nouveaux confirmés', ['nconf']), b: P('surv', 'Surveillance : nouveaux confirmés', g('cas', 'conf', VD)) }
  ];

  /* ---------- Évaluation ---------- */
  function valeursRapport(p, r) {
    var V = {}; A.ids(p).forEach(function (c) { var x = A.valeur(r, c.id); if (x != null) V[c.id] = x; }); return V;
  }
  function lecteur(V) { return function (id) { return V[id] == null ? null : V[id]; }; }

  /* Contrôles internes d'un rapport : [{ regle, ok, a, b }] (règles applicables seulement) */
  function controlerRapport(r) {
    var p = D.pilier(r.pilier); if (!p || A.estDepart(r)) return [];
    var V = valeursRapport(p, r), v = lecteur(V), out = [];
    (INTERNES[p.id] || []).forEach(function (R) {
      var x; try { x = R.f(v, r); } catch (e) { x = [null]; }
      if (x[0] == null) return;
      out.push({ regle: R, ok: x[0], a: x[1], b: x[2] });
    });
    return out;
  }
  /* Continuité : rapport r comparé au rapport de la veille (même clé, date − 1) */
  function controlerContinuite(r, tous) {
    var p = D.pilier(r.pilier), regles = CONTINUITE[r.pilier]; if (!p || !regles || A.estDepart(r)) return [];
    var veille = trouverVeille(r, tous); if (!veille) return [];
    var V0 = valeursRapport(p, veille), K0 = A.calculer(p, V0), V = valeursRapport(p, r), out = [];
    regles.forEach(function (R) {
      var attendu = R.veille(lecteur(V0), K0), x = V[R.champ];
      if (attendu == null || x == null) return;
      out.push({ regle: R, ok: eg(x, attendu), a: x, b: attendu, veille: veille });
    });
    return out;
  }
  function trouverVeille(r, tous) {
    var hier = A.decaler(r.date, -1);
    var cle = A.cleRapport(Object.assign({}, r, { date: hier }));
    return A.dedoublonner(tous).filter(function (x) { return A.cleRapport(x) === cle; })[0] || null;
  }
  /* Valeurs proposées pour les reports, d'après le rapport de la veille */
  function reportsProposes(r, tous) {
    var p = D.pilier(r.pilier), regles = CONTINUITE[r.pilier]; if (!p || !regles || A.estDepart(r)) return null;
    var veille = trouverVeille(r, tous); if (!veille) return null;
    var V0 = valeursRapport(p, veille), K0 = A.calculer(p, V0), o = {};
    regles.forEach(function (R) { var x = R.veille(lecteur(V0), K0); if (x != null) o[R.champ] = x; });
    return Object.keys(o).length ? { valeurs: o, veille: veille } : null;
  }

  /* Concordance entre piliers sur la période f, par zone et pour le total.
     → [{ regle, zone, a, b, statut: 'ok' | 'ecart' | 'manque', manque: [piliers sans rapport] }] */
  function concordances(rapports, f) {
    var C = {};
    D.PILIERS.forEach(function (p) { C[p.id] = A.consolider(rapports, p, f); });
    var zones = D.ZONES.filter(function (z) { return !f.zones || !f.zones.length || f.zones.indexOf(z) >= 0; });
    function cote(c, zone) {
      var parts = c.parts || [[c.pilier, c.ids]], val = null, manque = [];
      parts.forEach(function (pt) {
        var cz = zone === 'Total' ? (C[pt[0]].ordre.length ? C[pt[0]].total : null) : C[pt[0]].zones[zone];
        if (!cz) { manque.push(pt[0]); return; }
        // Pilier ayant rapporté : une case vide compte pour 0
        val = (val || 0) + (S(lecteur(cz.valeurs), pt[1]) || 0);
      });
      if (manque.length === parts.length) val = null;
      return { val: val, manque: manque };
    }
    var out = [];
    function statut(R, a, b) {
      if (a.val != null && b.val != null) return (R.type === 'eg' ? eg(a.val, b.val) : le(a.val, b.val)) ? 'ok' : 'ecart';
      // Un seul des deux piliers a rapporté : « rapport manquant » seulement si l'autre côté aurait dû avoir des chiffres
      if (R.type === 'le') return b.val == null && a.val > 0 ? 'manque' : null;
      return (a.val || b.val) ? 'manque' : null;
    }
    CONCORDANCES.forEach(function (R) {
      if (R.province) {
        var a0 = cote(R.a, 'Total'), b0 = cote(R.b, 'Total'), s0 = statut(R, a0, b0);
        if (s0) out.push({ regle: R, zone: 'Total', a: a0.val, b: b0.val, statut: s0, manque: a0.manque.concat(b0.manque) });
        return;
      }
      var ta = null, tb = null, n = 0;
      zones.forEach(function (z) {
        var a = cote(R.a, z), b = cote(R.b, z), st = statut(R, a, b);
        if (!st) return;
        out.push({ regle: R, zone: z, a: a.val, b: b.val, statut: st, manque: a.manque.concat(b.manque) });
        if (st !== 'manque') { ta = (ta || 0) + a.val; tb = (tb || 0) + b.val; n++; }
      });
      // Total : seulement les zones où les deux piliers ont rapporté
      if (zones.length > 1 && n > 1) out.push({ regle: R, zone: 'Total', a: ta, b: tb, statut: (R.type === 'eg' ? eg(ta, tb) : le(ta, tb)) ? 'ok' : 'ecart', manque: [], partiel: true });
    });
    return out;
  }

  /* Synthèse pour l'écran Qualité */
  function inspecter(rapports, f) {
    var retenus = A.dedoublonner(rapports).filter(function (r) {
      return !A.estDepart(r) && r.date >= f.debut && r.date <= f.fin && (!f.zones || !f.zones.length || f.zones.indexOf(r.zone) >= 0) && (!f.pilier || r.pilier === f.pilier);
    });
    var parRapport = retenus.map(function (r) {
      var I = controlerRapport(r).filter(function (x) { return !x.ok; }), T = controlerContinuite(r, rapports).filter(function (x) { return !x.ok; });
      return { r: r, internes: I, continuite: T, erreurs: I.filter(function (x) { return x.regle.grav === 'erreur'; }).length };
    });
    var conc = concordances(rapports, f).filter(function (x) { return !f.pilier || x.regle.a.pilier === f.pilier || x.regle.b.pilier === f.pilier || [].concat(x.regle.a.pilier).indexOf(f.pilier) >= 0; });
    var aCorriger = parRapport.filter(function (x) { return x.internes.length || x.continuite.length; })
      .sort(function (a, b) { return b.erreurs - a.erreurs || b.r.date.localeCompare(a.r.date); });
    return {
      rapports: retenus.length, sansAnomalie: parRapport.filter(function (x) { return !x.internes.length && !x.continuite.length; }).length,
      erreurs: parRapport.reduce(function (t, x) { return t + x.erreurs; }, 0),
      alertes: parRapport.reduce(function (t, x) { return t + x.internes.length - x.erreurs; }, 0),
      ruptures: parRapport.reduce(function (t, x) { return t + x.continuite.length; }, 0),
      concordances: conc, ecarts: conc.filter(function (x) { return x.statut === 'ecart'; }).length,
      manques: conc.filter(function (x) { return x.statut === 'manque'; }).length, concOk: conc.filter(function (x) { return x.statut === 'ok'; }).length,
      aCorriger: aCorriger
    };
  }

  var api = { INTERNES: INTERNES, CONTINUITE: CONTINUITE, CONCORDANCES: CONCORDANCES, controlerRapport: controlerRapport,
    controlerContinuite: controlerContinuite, reportsProposes: reportsProposes, concordances: concordances, inspecter: inspecter, nb: nb };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.QUALITE_PILIERS = api;
})(this);
