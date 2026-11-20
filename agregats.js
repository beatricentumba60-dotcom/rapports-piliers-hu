/* Calculs des rapports des piliers : valeurs d'un rapport, consolidation par zone et par période, complétude.
   Rapport : { id, pilier, date: 'AAAA-MM-JJ', zone, auteur, maj: ISO, donnees: { valeurs, listes, textes, entete, supprime } } */
(function (root) {
  'use strict';
  var D = root.PILIERS_DEF || (typeof require !== 'undefined' ? require('./piliers-def.js') : null);

  function num(x) { if (x === '' || x == null) return null; var n = Number(String(x).replace(',', '.')); return isNaN(n) ? null : n; }

  /* Valeur d'un champ dans UN rapport (null si non renseigné). Champs de liste : somme des lignes ;
     « sec.#col » = nombre de lignes où col vaut Oui ; « sec.# » = nombre de lignes renseignées. */
  function valeur(rap, id) {
    var d = rap.donnees || {};
    if (d.valeurs && d.valeurs[id] != null && d.valeurs[id] !== '') return num(d.valeurs[id]);
    var m = /^([^.]+)\.(#?)([^.]*)$/.exec(id);
    if (m && d.listes && d.listes[m[1]]) {
      var L = d.listes[m[1]].filter(function (x) { return x && Object.keys(x).some(function (k) { return x[k] != null && x[k] !== ''; }); });
      if (!L.length) return null;
      if (m[2] === '#') return m[3] ? L.filter(function (x) { return x[m[3]] === 'Oui'; }).length : L.length;
      var t = null; L.forEach(function (x) { var v = num(x[m[3]]); if (v != null) t = (t || 0) + v; });
      return t;
    }
    return null;
  }
  // Identifiants à consolider : champs numériques + comptes de liste utilisés par les calculs
  function ids(p) {
    var out = D.champs(p).map(function (c) { return { id: c.id, agg: c.agg }; });
    p.sections.forEach(function (sec) {
      if (sec.liste) { out.push({ id: sec.id + '.#', agg: 'sum' }); sec.liste.colonnes.forEach(function (c) { if (c[2] === 'choix') out.push({ id: sec.id + '.#' + c[0], agg: 'sum' }); }); }
    });
    return out;
  }
  // Clé d'unicité : un seul rapport par pilier, date, zone et unité (CTE, laboratoire…) ; le plus récent gagne
  function cleRapport(r) {
    var e = (r.donnees && r.donnees.entete) || {};
    return [r.pilier, r.date, r.zone].concat(Object.keys(e).sort().map(function (k) { return k + '=' + String(e[k] || '').trim().toLowerCase(); })).join('|');
  }
  function dedoublonner(rapports) {
    var par = {};
    rapports.forEach(function (r) {
      if (!r || !r.pilier || !r.date || !r.zone) return;
      var k = cleRapport(r);
      if (!par[k] || String(r.maj) > String(par[k].maj)) par[k] = r;
    });
    return Object.keys(par).map(function (k) { return par[k]; }).filter(function (r) { return !(r.donnees && r.donnees.supprime); });
  }

  /* Consolidation d'un pilier.
     f = { debut, fin (AAAA-MM-JJ), zones: [] (vide = toutes) }
     → { zones: { zone: { valeurs, rapports, dates } }, total: { valeurs }, ordre: [zones avec rapport] }
     agg 'sum' : somme sur les jours ; 'stock' : valeur du dernier jour renseigné. Entre unités d'un même jour et entre zones : somme. */
  function consolider(rapports, p, f) {
    var I = ids(p), Z = {};
    var garde = dedoublonner(rapports).filter(function (r) {
      return r.pilier === p.id && r.date >= f.debut && r.date <= f.fin && (!f.zones || !f.zones.length || f.zones.indexOf(r.zone) >= 0);
    });
    garde.forEach(function (r) {
      var z = Z[r.zone] || (Z[r.zone] = { jours: {}, rapports: [] });
      z.rapports.push(r);
      var j = z.jours[r.date] || (z.jours[r.date] = {});
      I.forEach(function (c) { var v = valeur(r, c.id); if (v != null) j[c.id] = (j[c.id] || 0) + v; });
    });
    var out = { zones: {}, total: { valeurs: {} }, ordre: [] };
    Object.keys(Z).forEach(function (zone) {
      var z = Z[zone], dates = Object.keys(z.jours).sort(), V = {};
      I.forEach(function (c) {
        if (c.agg === 'stock') {
          for (var i = dates.length - 1; i >= 0; i--) if (z.jours[dates[i]][c.id] != null) { V[c.id] = z.jours[dates[i]][c.id]; break; }
        } else dates.forEach(function (d) { var v = z.jours[d][c.id]; if (v != null) V[c.id] = (V[c.id] || 0) + v; });
      });
      out.zones[zone] = { valeurs: V, rapports: z.rapports, dates: dates };
      Object.keys(V).forEach(function (k) { out.total.valeurs[k] = (out.total.valeurs[k] || 0) + V[k]; });
    });
    out.total.rapports = garde;
    var rang = D.zonesPour(p);
    out.ordre = Object.keys(out.zones).sort(function (a, b) {
      var ia = rang.indexOf(a), ib = rang.indexOf(b); return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib) || a.localeCompare(b, 'fr');
    });
    return out;
  }
  function calculer(p, valeurs) {
    var v = function (id) { return valeurs[id] == null ? null : valeurs[id]; }, o = {};
    (p.calculs || []).forEach(function (c) { var x = null; try { x = c.f(v); } catch (e) { x = null; } o[c.id] = x == null || isNaN(x) || !isFinite(x) ? null : x; });
    return o;
  }

  /* Colonnes « clés » du tableau de bord et des diapositives.
     Ligne clé d'une grille additive (vivants + décès, confirmés + suspects) : une seule colonne = somme de la ligne. */
  function colonnesCles(p) {
    var out = [];
    (p.calculs || []).filter(function (k) { return k.cle && k.avant; }).forEach(function (k) { out.push({ lib: k.lib, calc: k, pct: k.pct }); });
    p.sections.forEach(function (sec) {
      (sec.champs || []).forEach(function (c) { if (c[3]) out.push({ lib: c[1], ids: [c[0]], stock: c[2] === 'stock' }); });
      if (sec.grille) sec.grille.lignes.forEach(function (l) {
        if (!l[3]) return;
        if (sec.grille.somme) out.push({ lib: l[1], ids: sec.grille.colonnes.map(function (c) { return sec.id + '.' + l[0] + '.' + c[0]; }), stock: l[2] === 'stock' });
        else sec.grille.colonnes.forEach(function (c) { out.push({ lib: l[1] + ' – ' + c[1], ids: [sec.id + '.' + l[0] + '.' + c[0]], stock: l[2] === 'stock' }); });
      });
    });
    (p.calculs || []).filter(function (k) { return k.cle && !k.avant; }).forEach(function (k) { out.push({ lib: k.lib, calc: k, pct: k.pct }); });
    return out;
  }
  function valCol(col, V, K) {
    if (col.calc) return K[col.calc.id] == null ? null : K[col.calc.id];
    var t = null; col.ids.forEach(function (id) { if (V[id] != null) t = (t || 0) + V[id]; }); return t;
  }

  /* Complétude d'une date : { pilier: { zone: nombre de rapports } } */
  function completude(rapports, date) {
    var o = {};
    dedoublonner(rapports).forEach(function (r) {
      if (r.date !== date) return;
      var p = o[r.pilier] || (o[r.pilier] = {}); p[r.zone] = (p[r.zone] || 0) + 1;
    });
    return o;
  }

  /* Textes et listes des rapports retenus (pour l'écran, Excel et les diapositives) */
  function textes(rapports, id) {
    return rapports.map(function (r) { return { zone: r.zone, date: r.date, r: r, t: ((r.donnees || {}).textes || {})[id] }; })
      .filter(function (x) { return x.t && String(x.t).trim(); });
  }
  function lignesListe(rapports, sec) {
    var out = [];
    rapports.forEach(function (r) {
      ((r.donnees || {}).listes || {})[sec] && r.donnees.listes[sec].forEach(function (l) {
        if (l && Object.keys(l).some(function (k) { return l[k] != null && l[k] !== ''; })) out.push({ zone: r.zone, date: r.date, r: r, l: l });
      });
    });
    return out;
  }

  // Dates : texte AAAA-MM-JJ (heure locale)
  function jour(d) { d = d || new Date(); return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
  function decaler(j, n) { var p = j.split('-'); var d = new Date(+p[0], +p[1] - 1, +p[2] + n); return jour(d); }
  function fr(j) { if (!j) return ''; var p = j.split('-'); return p[2] + '/' + p[1] + '/' + p[0]; }

  var api = { colonnesCles: colonnesCles, valCol: valCol, valeur: valeur, ids: ids, cleRapport: cleRapport, dedoublonner: dedoublonner, consolider: consolider, calculer: calculer,
    completude: completude, textes: textes, lignesListe: lignesListe, jour: jour, decaler: decaler, fr: fr, num: num };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.AGREGATS = api;
})(this);
