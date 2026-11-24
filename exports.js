/* Exports de « Rapports des piliers » : classeur Excel (SheetJS) et diapositives PowerPoint (PptxGenJS). */
(function (root) {
  'use strict';
  var D = root.PILIERS_DEF, A = root.AGREGATS;

  function filtrer(rapports, f, pilierId) {
    return A.dedoublonner(rapports).filter(function (r) {
      if (A.estDepart(r) && !f.cumul) return false; // valeurs de départ : seulement dans le cumul
      return (!pilierId || r.pilier === pilierId) && r.date >= f.debut && r.date <= f.fin && (!f.zones.length || f.zones.indexOf(r.zone) >= 0);
    }).sort(function (a, b) { return a.date.localeCompare(b.date) || a.zone.localeCompare(b.zone, 'fr'); });
  }
  function libPeriode(f) {
    return f.periode === '0' ? 'du ' + A.fr(f.fin) : f.periode === 'tout' ? 'cumul jusqu’au ' + A.fr(f.fin) : 'du ' + A.fr(f.debut) + ' au ' + A.fr(f.fin);
  }
  function sectionsTexte(p) { var o = []; D.sections(p).forEach(function (s) { (s.textes || []).forEach(function (t) { o.push(t); }); }); return o; }

  /* ---------- Excel ---------- */
  function excel(rapports, f, liens) {
    var X = root.XLSX, wb = X.utils.book_new();
    var noms = {};
    function feuille(nom, aoa, largeurs) {
      var n = nom.replace(/\s*\/\s*/g, ' - ').replace(/[\\?*\[\]:]/g, '-').slice(0, 31); var i = 2; while (noms[n]) n = nom.slice(0, 28) + ' ' + i++; noms[n] = 1;
      var ws = X.utils.aoa_to_sheet(aoa);
      if (largeurs) ws['!cols'] = largeurs.map(function (w) { return { wch: w }; });
      X.utils.book_append_sheet(wb, ws, n);
      return ws;
    }
    function pct(ws, r, c) { var a = X.utils.encode_cell({ r: r, c: c }); if (ws[a] && typeof ws[a].v === 'number') ws[a].z = '0.0%'; }
    // Lisez-moi
    feuille('Lisez-moi', [['TASK FORCE PRÉSIDENTIELLE EBOLA 17'], ['MINISTÈRE DE LA SANTÉ PUBLIQUE, HYGIÈNE ET PRÉVOYANCE SOCIALE'],
      ['SYSTÈME DE GESTION DE L’INCIDENT MVE 17'], ['17e Epidémie de la Maladie à Virus Ebola – Division Provinciale de la santé du Haut-Uele'], [], ['Rapports quotidiens des piliers'], [],
      ['Période', libPeriode(f)], ['Zone', f.zone || 'Toutes les zones'], ['Pilier', f.pilier ? D.pilier(f.pilier).nom : 'Tous'],
      ['Exporté le', new Date().toLocaleString('fr-FR')], [],
      ['Feuille « Synthèse » : consolidation par pilier et par zone. Les valeurs « fin de journée » (reports, hospitalisés, stocks…) prennent le dernier jour de la période ; les autres sont additionnées.'],
      ['Une feuille par pilier : un rapport par ligne (une ligne par zone, date et unité). En cumul, les lignes « Départ » sont les valeurs de départ (cumul avant la mise en service).'],
      ['Feuilles « Défis », « Détail par PoE-PoC » et « Feedbacks CREC » : les lignes saisies dans ces tableaux.']], [22, 60]);
    // Complétude
    var cols = D.ZONES.concat([D.PROVINCE]), C = A.completude(rapports, f.fin);
    feuille('Complétude ' + f.fin, [['Pilier'].concat(cols, ['ZS ayant rapporté'])].concat(D.PILIERS.map(function (p) {
      var c = C[p.id] || {}; return [p.nom].concat(cols.map(function (z) { return c[z] || 0; }), [D.ZONES.filter(function (z) { return c[z]; }).length]);
    })), [38].concat(cols.map(function () { return 9; }), [10]));
    var piliers = D.PILIERS.filter(function (p) { return !f.pilier || p.id === f.pilier; });
    // Synthèse
    var syn = [], pcts = [];
    piliers.forEach(function (p) {
      var c = A.consolider(rapports, p, f); if (!c.ordre.length) return;
      var ch = D.champs(p), ks = p.calculs || [];
      syn.push([p.nom + ' – ' + libPeriode(f)]);
      syn.push(['Zone'].concat(ch.map(function (x) { return x.sec.titre + ' : ' + x.lib + (x.agg === 'stock' ? ' (fin de journée)' : ''); }), ks.map(function (k) { return k.lib; })));
      var ligne = function (nom, V) {
        var K = A.calculer(p, V);
        ks.forEach(function (k, j) { if (k.pct) pcts.push([syn.length, 1 + ch.length + j]); });
        return [nom].concat(ch.map(function (x) { return V[x.id] == null ? null : V[x.id]; }), ks.map(function (k) { return K[k.id]; }));
      };
      c.ordre.forEach(function (z) { syn.push(ligne(z, c.zones[z].valeurs)); });
      if (c.ordre.length > 1) syn.push(ligne('TOTAL', c.total.valeurs));
      syn.push([]);
    });
    if (syn.length) { var wsS = feuille('Synthèse', syn, [26]); pcts.forEach(function (x) { pct(wsS, x[0], x[1]); }); }
    // Une feuille par pilier
    piliers.forEach(function (p) {
      var L = filtrer(rapports, f, p.id); if (!L.length) return;
      var ch = D.champs(p), ks = p.calculs || [], tx = sectionsTexte(p), en = p.entete || [];
      var aoa = [['Date', 'Zone'].concat(en.map(function (e) { return e[1]; }), ['Auteur', 'Modifié le'],
        ch.map(function (x) { return x.sec.titre + ' : ' + x.lib; }), ks.map(function (k) { return k.lib; }), tx.map(function (t) { return t[1]; }))];
      L.forEach(function (r) {
        var V = {}; A.ids(p).forEach(function (c) { var v = A.valeur(r, c.id); if (v != null) V[c.id] = v; });
        var K = A.calculer(p, V), d = r.donnees || {};
        aoa.push([(A.estDepart(r) ? 'Départ (cumul au ' + A.fr(r.date) + ')' : A.fr(r.date)), r.zone].concat(en.map(function (e) { return (d.entete || {})[e[0]] || ''; }), [r.auteur || '', r.maj ? new Date(r.maj).toLocaleString('fr-FR') : ''],
          ch.map(function (x) { return V[x.id] == null ? null : V[x.id]; }), ks.map(function (k) { return K[k.id]; }), tx.map(function (t) { return (d.textes || {})[t[0]] || ''; })));
      });
      var ws = feuille(p.court, aoa, [11, 16].concat(en.map(function () { return 16; }), [22, 18]));
      var c0 = 4 + en.length + ch.length;
      ks.forEach(function (k, j) { if (k.pct) for (var i = 1; i < aoa.length; i++) pct(ws, i, c0 + j); });
    });
    // Défis, PoE-PoC, feedbacks
    var defis = [['Pilier', 'Date', 'Zone', 'Défi / problème', 'Piste de solution', 'Responsable', 'Délai']];
    piliers.forEach(function (p) { A.lignesListe(filtrer(rapports, f, p.id), 'defis').forEach(function (x) { defis.push([p.nom, A.fr(x.date), x.zone, x.l.defi || '', x.l.solution || '', x.l.resp || '', x.l.delai || '']); }); });
    if (defis.length > 1) feuille('Défis', defis, [30, 11, 16, 50, 50, 20, 14]);
    var poe = D.pilier('poe'), sp = poe.sections[0].liste.colonnes;
    if (!f.pilier || f.pilier === 'poe') {
      var pts = [['Date', 'Zone'].concat(sp.map(function (c) { return c[1]; }))];
      A.lignesListe(filtrer(rapports, f, 'poe'), 'pts').forEach(function (x) {
        pts.push([A.fr(x.date), x.zone].concat(sp.map(function (c) { return c[2] === 'n' ? A.num(x.l[c[0]]) : (x.l[c[0]] || ''); })));
      });
      if (pts.length > 1) feuille('Détail par PoE-PoC', pts, [11, 16, 24, 12]);
    }
    if (!f.pilier || f.pilier === 'crec') {
      var fb = [['Date', 'Zone', 'Feedback', 'Source et localisation', 'Type', 'Pilier concerné']];
      A.lignesListe(filtrer(rapports, f, 'crec'), 'fb').forEach(function (x) { fb.push([A.fr(x.date), x.zone, x.l.texte || '', x.l.source || '', x.l.type || '', x.l.pilier || '']); });
      if (fb.length > 1) feuille('Feedbacks CREC', fb, [11, 16, 70, 26, 14, 16]);
    }
    // Valeurs saisies par aire de santé (facultatif)
    var pa = [['Pilier', 'Date', 'Zone de santé', 'Aire de santé', 'Population 2026', 'Indicateur', 'Valeur']];
    piliers.forEach(function (p) {
      var libs = {}; D.champs(p).forEach(function (c) { libs[c.id] = c.lib; });
      filtrer(rapports, f, p.id).forEach(function (r) {
        var AI = (r.donnees || {}).aires || {};
        D.aires(r.zone).filter(function (a) { return AI[a]; }).forEach(function (a) {
          Object.keys(AI[a]).forEach(function (id) { if (AI[a][id] != null) pa.push([p.nom, A.fr(r.date), r.zone, a, D.populationAire(r.zone, a), libs[id] || id, AI[a][id]]); });
        });
      });
    });
    if (pa.length > 1) feuille('Par aire de santé', pa, [30, 11, 16, 22, 14, 50, 10]);
    // Photos des activités : légende et lien Google Drive
    var ph = [['Pilier', 'Date', 'Zone', 'Légende', 'Auteur', 'Lien Google Drive']];
    piliers.forEach(function (p) {
      filtrer(rapports, f, p.id).forEach(function (r) {
        ((r.donnees || {}).photos || []).forEach(function (x) { ph.push([p.nom, A.fr(r.date), r.zone, x.legende || '', r.auteur || '', (liens && liens[x.id]) || 'pas encore envoyée']); });
      });
    });
    if (ph.length > 1) feuille('Photos', ph, [30, 11, 16, 50, 24, 50]);
    // Qualité : incohérences, ruptures avec la veille, concordance entre piliers
    var Q = root.QUALITE_PILIERS;
    if (Q) {
      var R = Q.inspecter(rapports, f), nomP = function (id) { return [].concat(id).map(function (x) { var p = D.pilier(x); return p ? p.court : x; }).join(' + '); };
      var qa = [['Type', 'Gravité', 'Contrôle', 'Pilier(s)', 'Zone', 'Date', 'Valeur 1', 'Valeur 2', 'Statut', 'Auteur']];
      R.aCorriger.forEach(function (x) {
        x.internes.forEach(function (i) { qa.push(['Cohérence du rapport', i.regle.grav === 'erreur' ? 'Incohérence' : 'À vérifier', i.regle.lib, nomP(x.r.pilier), x.r.zone, A.fr(x.r.date), i.a, i.b, 'À corriger', x.r.auteur || '']); });
        x.continuite.forEach(function (c) { qa.push(['Continuité avec la veille', 'À vérifier', c.regle.lib, nomP(x.r.pilier), x.r.zone, A.fr(x.r.date), c.a, c.b, 'Saisi / veille', x.r.auteur || '']); });
      });
      R.concordances.forEach(function (c) {
        qa.push(['Concordance entre piliers', c.statut === 'ecart' ? 'Écart' : c.statut === 'ok' ? 'Concordant' : 'Rapport manquant', c.regle.lib + ' (' + c.regle.a.lib + ' / ' + c.regle.b.lib + ')',
          nomP(c.regle.a.pilier) + ' ↔ ' + nomP(c.regle.b.pilier), c.zone, libPeriode(f), c.a, c.b,
          c.statut === 'ok' ? 'OK' : c.statut === 'ecart' ? 'Écart à rapprocher' : 'Pas de rapport : ' + nomP(c.manque), '']);
      });
      feuille('Qualité', qa, [24, 14, 70, 26, 16, 22, 10, 10, 26, 22]);
    }
    X.writeFile(wb, (f.pilier ? 'Rapport_' + sansAccents(D.pilier(f.pilier).court).replace(/[\s\/]+/g, '_') + '_HU_' : 'Rapports_piliers_HU_') + f.fin + (f.zone ? '_' + f.zone.replace(/\s+/g, '_') : '') + '.xlsx');
  }

  /* ---------- PowerPoint ---------- */
  var W = 13.333, H = 7.5, ENCRE = '0F3D3E', GRIS = '5D6B78';
  function sansAccents(t) { return String(t).normalize('NFD').replace(/[\u0300-\u036f]/g, ''); }
  function morceaux(arr, n) { var o = []; for (var i = 0; i < arr.length; i += n) o.push(arr.slice(i, i + n)); return o.length ? o : [[]]; }
  function f1(v, pct) { if (v == null) return '–'; return pct ? (Math.round(v * 1000) / 10).toString().replace('.', ',') + ' %' : (Math.round(v * 10) / 10).toLocaleString('fr-FR'); }
  // Emblème (embleme.png) en data URL pour l'intégrer à la présentation
  function chargerEmbleme(nom) {
    return fetch('' + nom).then(function (r) { if (!r.ok) throw 0; return r.blob(); }).then(function (b) {
      return new Promise(function (res) { var fr = new FileReader(); fr.onload = function () { res(fr.result); }; fr.onerror = function () { res(null); }; fr.readAsDataURL(b); });
    }).catch(function () { return null; });
  }
  function diapos(rapports, f, chargerPhoto) {
    // Photos des rapports retenus (60 au plus), chargées avant la construction de la présentation
    var ids = [];
    filtrer(rapports, f).forEach(function (r) { if (!f.pilier || r.pilier === f.pilier) ((r.donnees || {}).photos || []).forEach(function (x) { ids.push(x.id); }); });
    ids = ids.slice(0, 60);
    var photos = {};
    var pr = chargerPhoto ? ids.reduce(function (p, id) { return p.then(function () { return chargerPhoto(id).then(function (d) { if (d) photos[id] = d; }, function () { /* absente */ }); }); }, Promise.resolve()) : Promise.resolve();
    return pr.then(function () { return Promise.all([chargerEmbleme('entete-officiel.png'), chargerEmbleme('embleme-petit.png'), chargerEmbleme('diapo-merci.jpg')]); })
      .then(function (e) { return diapos_(rapports, f, e[0], e[1], photos, e[2]); });
  }
  // Couleurs de l'en-tête officiel du SGI MVE 17
  var OFF = { marine: '1F2A78', or: 'C9A227', rouge: 'E30613', vert: '00A651', bleu: '1D3F9A', rougeL: 'D52B1E', jaune: 'F2B705' };
  function diapos_(rapports, f, banniere, embPetit, photos, merci) {
    photos = photos || {};
    var P = new root.PptxGenJS();
    P.layout = 'LAYOUT_WIDE'; P.author = 'DPS Haut-Uélé – SGI MVE'; P.company = 'SGI MVE 17 – DPS Haut-Uélé'; P.title = 'Rapports des piliers ' + libPeriode(f);
    // Seulement la date (et la zone si un filtre de zone est choisi)
    var periode = f.periode === '0' ? A.fr(f.fin) : libPeriode(f).replace(/^du /, 'Du '), lieu = f.zone || '';
    var etiquette = lieu ? lieu + ' · ' + periode : periode;
    function nouvelle(titre, couleur, sousTitre) {
      var s = P.addSlide();
      s.background = { color: 'FFFFFF' };
      s.addShape(P.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.85, fill: { color: couleur || ENCRE }, line: { color: couleur || ENCRE } });
      s.addText(titre, { x: 0.4, y: 0.1, w: W - 3.2, h: 0.65, fontFace: 'Arial', fontSize: 22, bold: true, color: 'FFFFFF', valign: 'middle', fit: 'shrink' });
      s.addText(lieu ? lieu + '\n' + periode : periode, { x: W - 2.9, y: 0.1, w: 2.6, h: 0.65, fontFace: 'Arial', fontSize: 10, color: 'FFFFFF', align: 'right', valign: 'middle' });
      if (sousTitre) s.addText(sousTitre, { x: 0.4, y: 0.95, w: W - 0.8, h: 0.4, fontFace: 'Arial', fontSize: 13, color: GRIS, italic: true });
      // Filet tricolore sous le bandeau, comme l'en-tête officiel
      [OFF.bleu, OFF.rougeL, OFF.jaune].forEach(function (c, i) { s.addShape(P.ShapeType.rect, { x: i * W / 3, y: 0.85, w: W / 3, h: 0.05, fill: { color: c }, line: { color: c } }); });
      if (embPetit) s.addImage({ data: embPetit, x: W - 0.75, y: H - 0.62, w: 0.48, h: 0.45 });
      s.addText('Task Force Présidentielle Ebola 17 · SGI MVE 17 · Division Provinciale de la santé du Haut-Uele', { x: 0.4, y: H - 0.4, w: W - 1.4, h: 0.3, fontFace: 'Arial', fontSize: 9, color: '8A97A3' });
      return s;
    }
    function tableau(s, entete, lignes, opts) {
      opts = opts || {};
      var nb = entete.length, wl = opts.wl || Math.min(3.2, (W - 0.8) * 0.3), wr = nb > 1 ? (W - 0.8 - wl) / (nb - 1) : 0;
      var fs = opts.fs || (nb > 10 ? 9 : nb > 8 ? 10 : nb > 5 ? 12 : 14);
      var rows = [entete.map(function (h, i) { return { text: String(h), options: { bold: true, color: 'FFFFFF', fill: { color: opts.couleur || ENCRE }, align: i ? 'center' : 'left', valign: 'middle' } }; })]
        .concat(lignes.map(function (l, j) {
          var tot = /^total$/i.test(String(l[0]));
          return l.map(function (v, i) { return { text: v == null ? '' : String(v), options: { align: i && !opts.gauche ? 'right' : 'left', bold: tot, fill: { color: tot ? 'E9EEF0' : (j % 2 ? 'F6F8F9' : 'FFFFFF') } } }; });
        }));
      s.addTable(rows, { x: 0.4, y: opts.y || 1.45, w: W - 0.8, colW: opts.colW || [wl].concat(Array(Math.max(0, nb - 1)).fill(wr)), fontFace: 'Arial', fontSize: fs,
        border: { type: 'solid', pt: 0.5, color: 'D5DCE1' }, autoPage: true, autoPageRepeatHeader: true, autoPageLineWeight: 0, autoPageCharWeight: 0, newSlideStartY: 1.0 });
    }
    function puces(s, items, y, h, x, w) {
      var runs = [];
      items.forEach(function (it) {
        runs.push({ text: it[0], options: { bold: true, bullet: true, breakLine: false } });
        runs.push({ text: it[1], options: { breakLine: true } });
      });
      s.addText(runs.length ? runs : [{ text: 'RAS' }], { x: x || 0.5, y: y || 1.45, w: w || W - 1, h: h || H - 2.1, fontFace: 'Arial', fontSize: 13, color: '16212B', valign: 'top', paraSpaceAfter: 6, fit: 'shrink' });
    }
    function textesEnPuces(L, multiJours) {
      return L.map(function (t) { return [t.zone + (multiJours ? ' (' + A.fr(t.date) + ')' : '') + ' : ', String(t.t).replace(/\s*\n\s*/g, ' ; ')]; });
    }
    var multi = f.debut !== f.fin;
    /* ---------- Graphiques ---------- */
    var PAL = ['1B5E60', 'B97700', '2E86AB', '7A4E9E', '5D6B78', '23804F'];
    function couleurs(coul, n) { return [coul].concat(PAL.filter(function (c) { return c !== coul; })).slice(0, Math.max(1, n)); }
    function z0(v) { return v == null ? 0 : Math.round(v * 10) / 10; }
    function court(t, n) { t = String(t); return t.length > n ? t.slice(0, n - 1) + '…' : t; }
    // Barres (verticales par défaut). series : [{ name, labels, values }]
    function barres(s, series, coul, o) {
      o = o || {};
      var n = series.length, nCat = (series[0] || { labels: [] }).labels.length;
      s.addChart(P.ChartType.bar, series, {
        x: o.x || 0.4, y: o.y || 1.45, w: o.w || W - 0.8, h: o.h || H - 2.05,
        barDir: o.horiz ? 'bar' : 'col', barGrouping: 'clustered', barGapWidthPct: n > 1 ? 40 : 70,
        chartColors: n > 1 ? couleurs(coul, n) : [coul],
        catAxisLabelFontFace: 'Arial', catAxisLabelFontSize: o.fsCat || (nCat > 10 ? 9 : 11), catAxisLabelColor: '33414D', catAxisOrientation: o.horiz ? 'maxMin' : 'minMax',
        valAxisLabelFontFace: 'Arial', valAxisLabelFontSize: 9, valAxisLabelColor: GRIS, valAxisLabelFormatCode: o.pct ? '0"%"' : '#,##0',
        valGridLine: { color: 'E3E8EB', size: 0.5 }, catGridLine: { style: 'none' },
        showValue: true, dataLabelFontFace: 'Arial', dataLabelFontSize: n > 2 ? 8 : 10, dataLabelColor: '16212B', dataLabelFormatCode: o.pct ? '0.0"%"' : '#,##0.#',
        dataLabelPosition: 'outEnd',
        showLegend: n > 1, legendPos: 'b', legendFontFace: 'Arial', legendFontSize: 11,
        showTitle: !!o.titre, title: o.titre || '', titleFontFace: 'Arial', titleFontSize: 13, titleColor: GRIS
      });
    }
    function courbes(s, series, coul, o) {
      o = o || {};
      s.addChart(P.ChartType.line, series, {
        x: o.x || 0.4, y: o.y || 1.45, w: o.w || W - 0.8, h: o.h || H - 2.05,
        chartColors: couleurs(coul, series.length), lineSize: 2.5, lineDataSymbol: 'circle', lineDataSymbolSize: 7,
        catAxisLabelFontFace: 'Arial', catAxisLabelFontSize: 10, catAxisLabelColor: '33414D',
        valAxisLabelFontFace: 'Arial', valAxisLabelFontSize: 9, valAxisLabelColor: GRIS, valAxisLabelFormatCode: '#,##0',
        valGridLine: { color: 'E3E8EB', size: 0.5 }, catGridLine: { style: 'none' },
        showValue: series.length === 1, dataLabelFontSize: 9, dataLabelPosition: 't',
        showLegend: true, legendPos: 'b', legendFontFace: 'Arial', legendFontSize: 11
      });
    }
    // Petits graphiques côte à côte, chacun avec sa propre échelle : [{ titre, labels, values }]
    function multiples(s, L, coul, o) {
      o = o || {}; var n = L.length; if (!n) return;
      var nc = n === 4 ? 2 : Math.min(n, 3), nr = Math.ceil(n / nc), y0 = o.y || 1.45, hT = H - 0.5 - y0, gap = 0.2;
      var w = (W - 0.8 - gap * (nc - 1)) / nc, h = (hT - gap * (nr - 1)) / nr;
      L.forEach(function (g, i) {
        var x = 0.4 + (i % nc) * (w + gap), y = y0 + Math.floor(i / nc) * (h + gap), ser = [{ name: g.titre, labels: g.labels, values: g.values }];
        var opt = { x: x, y: y, w: w, h: h, titre: court(g.titre, 60) };
        if (o.line) s.addChart(P.ChartType.line, ser, { x: x, y: y, w: w, h: h, chartColors: [couleurs(coul, 4)[i % 4]], lineSize: 2.5, lineDataSymbol: 'circle', lineDataSymbolSize: 6,
          catAxisLabelFontFace: 'Arial', catAxisLabelFontSize: 8, catAxisLabelColor: '33414D', valAxisLabelFontFace: 'Arial', valAxisLabelFontSize: 8, valAxisLabelColor: GRIS, valAxisLabelFormatCode: '#,##0', valAxisMinVal: 0,
          valGridLine: { color: 'E3E8EB', size: 0.5 }, catGridLine: { style: 'none' }, showValue: g.labels.length <= 14, dataLabelFontSize: 8, dataLabelPosition: 't',
          showLegend: false, showTitle: true, title: court(g.titre, 60), titleFontFace: 'Arial', titleFontSize: 12, titleColor: '33414D' });
        else barres(s, ser, couleurs(coul, 4)[o.uneCouleur ? 0 : i % 4], Object.assign(opt, { fsCat: g.labels.length > 8 ? 8 : 10, horiz: o.horiz }));
      });
    }
    // Tuiles de chiffres clés (6 par rangée)
    function tuiles(s, items, y0, coul) {
      morceaux(items, 6).forEach(function (rang, r) {
        // k = [libellé, valeur du jour, cumul (facultatif)]
        var cum = items.some(function (k) { return k.length > 2; }), th = cum ? 1.5 : 1.15, pas = th + 0.15;
        var bw = (W - 0.8 - 0.15 * (rang.length - 1)) / Math.max(1, rang.length), y = y0 + r * pas;
        rang.forEach(function (k, i) {
          var x = 0.4 + i * (bw + 0.15);
          s.addShape(P.ShapeType.roundRect, { x: x, y: y, w: bw, h: th, fill: { color: 'F1F4F5' }, line: { color: 'F1F4F5' }, rectRadius: 0.08 });
          s.addText(k[1], { x: x, y: y + 0.05, w: bw, h: 0.6, fontFace: 'Arial', fontSize: 24, bold: true, color: coul, align: 'center' });
          s.addText(k[0], { x: x + 0.05, y: y + 0.6, w: bw - 0.1, h: 0.5, fontFace: 'Arial', fontSize: 10, color: GRIS, align: 'center', valign: 'top', fit: 'shrink' });
          if (cum) {
            s.addShape(P.ShapeType.line, { x: x + 0.2, y: y + 1.1, w: bw - 0.4, h: 0, line: { color: 'D5DCE1', width: 0.75 } });
            s.addText(k[2] == null ? '' : [{ text: 'Cumul : ', options: { color: GRIS } }, { text: k[2], options: { bold: true, color: '16212B' } }],
              { x: x, y: y + 1.12, w: bw, h: 0.33, fontFace: 'Arial', fontSize: 12, align: 'center', valign: 'middle', fit: 'shrink' });
          }
        });
      });
      var cum0 = items.some(function (k) { return k.length > 2; });
      return y0 + Math.ceil(items.length / 6) * ((cum0 ? 1.5 : 1.15) + 0.15);
    }

    /* Activités réalisées : chaque activité avec ses photos sur la même diapositive */
    function photoSur(s, x, x0, y0, w, h) {
      s.addShape(P.ShapeType.rect, { x: x0, y: y0, w: w, h: h, fill: { color: 'F1F4F5' }, line: { color: 'F1F4F5' } });
      s.addImage({ data: x.data, x: x0, y: y0, w: w, h: h, sizing: { type: 'contain', w: w, h: h } });
      if (x.legende) s.addText(x.legende, { x: x0, y: y0 + h + 0.05, w: w, h: 0.4, fontFace: 'Arial', fontSize: 11, color: GRIS, italic: true, align: 'center', valign: 'top', fit: 'shrink' });
    }
    function activitesRealisees(p, coul, R, avecZone) {
      var act = [], libres = [];
      R.forEach(function (r) {
        var pre = avecZone ? r.zone + (multi ? ' (' + A.fr(r.date) + ')' : '') + ' : ' : '', mien = [];
        A.activites(((r.donnees || {}).textes || {}).realisees).forEach(function (a) { var o = { t: pre + a, brut: a, ph: [] }; act.push(o); mien.push(o); });
        ((r.donnees || {}).photos || []).forEach(function (x) {
          if (!photos[x.id]) return;
          var o = { data: photos[x.id], legende: [avecZone ? r.zone : '', x.legende || ''].filter(Boolean).join(' – ') }, cible = null;
          mien.forEach(function (a) { if (!cible && x.activite && a.brut === x.activite) cible = a; });
          (cible ? cible.ph : libres).push(o);
        });
      });
      var titre = p.court + ' – activités réalisées';
      if (!act.length && !libres.length) { puces(nouvelle(titre, coul), [], 1.5, H - 2.2); return; }
      // Dans l'ordre : activités sans photo regroupées en liste numérotée ; activité avec photos = sa diapositive
      var attente = [];
      function vider() { morceaux(attente, 10).forEach(function (part) { if (part.length) puces(nouvelle(titre, coul), part.map(function (x) { return [x.n + '. ', x.t]; }), 1.5, H - 2.2); }); attente = []; }
      act.forEach(function (a, i) {
        if (!a.ph.length) { attente.push({ n: i + 1, t: a.t }); return; }
        vider();
        morceaux(a.ph, 2).forEach(function (deux) {
          var s = nouvelle(titre, coul);
          s.addText([{ text: (i + 1) + '. ', options: { bold: true, color: coul } }, { text: a.t }], { x: 0.4, y: 1.0, w: W - 0.8, h: 0.75, fontFace: 'Arial', fontSize: 16, color: '16212B', valign: 'middle', fit: 'shrink' });
          if (deux.length === 1) photoSur(s, deux[0], (W - 8.2) / 2, 1.85, 8.2, 4.6);
          else deux.forEach(function (x, k) { photoSur(s, x, 0.4 + k * 6.37, 1.85, 6.15, 4.6); });
        });
      });
      vider();
      morceaux(libres, 2).forEach(function (deux) {
        if (!deux.length) return;
        var s = nouvelle(titre + ' – autres photos d’actions', coul);
        deux.forEach(function (x, k) { photoSur(s, x, 0.4 + k * 6.37, 1.2, 6.15, 5.1); });
      });
    }

    // Couverture : en-tête officiel (Task Force, Ministère, SGI MVE 17, DPS Haut-Uele)
    var c0 = P.addSlide();
    c0.background = { color: 'FFFFFF' };
    if (banniere) { var bw0 = W - 1.0, bh0 = bw0 * 402 / 1252; c0.addImage({ data: banniere, x: 0.5, y: 0.35, w: bw0, h: bh0 }); }
    var cf = 'Century Gothic';
    // Pilier X / zone de santé / date ; coordination provinciale : la date seule
    var nomZ = !f.zone ? '' : (D.ZONES.indexOf(f.zone) >= 0 ? 'ZS ' : '') + f.zone.toUpperCase();
    var lignesC = f.unRapport ? ['Pilier ' + D.pilier(f.pilier).nom, nomZ, A.fr(f.fin)].filter(Boolean)
      : (f.pilier ? ['Pilier ' + D.pilier(f.pilier).nom, periode] : [periode]);
    var hC = 0.25 + lignesC.length * 0.5, yC = 4.75;
    c0.addShape(P.ShapeType.rect, { x: 2.4, y: yC, w: W - 4.8, h: hC, fill: { color: 'F1F5F6' }, line: { color: 'F1F5F6' } });
    lignesC.forEach(function (t, i) {
      var titre = lignesC.length > 1 && i === 0;
      c0.addText(t, { x: 2.4, y: yC + 0.12 + i * 0.5, w: W - 4.8, h: 0.5, fontFace: cf, fontSize: titre ? 22 : 18, bold: true, color: titre ? ENCRE : (i === lignesC.length - 1 ? OFF.or : ENCRE), align: 'center', valign: 'middle' });
    });

    // Complétude (rapport de la coordination seulement)
    var C = A.completude(rapports, f.fin), cols = D.ZONES.concat([D.PROVINCE]);
    var sc = f.unRapport ? null : nouvelle('Complétude des rapports du ' + A.fr(f.fin));
    if (sc) tableau(sc, ['Pilier'].concat(cols.map(function (z) { return z === D.PROVINCE ? 'Prov.' : z.replace('Boma Mangbetu', 'Boma M.'); }), ['ZS']),
      D.PILIERS.filter(function (p) { return !f.pilier || p.id === f.pilier; }).map(function (p) { var c = C[p.id] || {}; return [p.court].concat(cols.map(function (z) { return c[z] ? '✓' : ''; }), [D.ZONES.filter(function (z) { return c[z]; }).length + '/13']); }),
      { wl: 2.4, fs: 10, gauche: false });

    var U = !!f.unRapport;
    D.PILIERS.filter(function (p) { return !f.pilier || p.id === f.pilier; }).forEach(function (p) {
      var c = A.consolider(rapports, p, f); if (!c.ordre.length) return;
      var coul = p.couleur.replace('#', ''), R = c.total.rapports, Kt = A.calculer(p, c.total.valeurs), MZ = c.ordre.length > 1;
      var avecZone = MZ || multi;
      var colsK = A.colonnesCles(p), valK = function (col, V) { return A.valCol(col, V, A.calculer(p, V)); };
      var nombres = colsK.filter(function (k) { return !k.pct; });

      // 1. Points saillants
      var s1 = nouvelle(p.court + ' – points saillants', coul, MZ ? c.ordre.length + ' zones ayant rapporté' : '');
      var ps = []; A.textes(R, 'saillants').forEach(function (t) { A.activites(t.t).forEach(function (l) { ps.push([avecZone ? t.zone + (multi ? ' (' + A.fr(t.date) + ')' : '') + ' : ' : '', l]); }); });
      puces(s1, ps, 1.5, H - 2.2);

      // 2. Activités réalisées (avec leurs photos)
      activitesRealisees(p, coul, R, avecZone);

      // 3. Indicateurs clés : tous les chiffres, puis les graphiques (aires de santé ou zones)
      // Cumul toujours présenté : depuis le début de l'épidémie, valeurs de départ comprises (sauf si la période est déjà le cumul)
      var cC = f.cumul ? null : A.consolider(rapports, p, { debut: '2000-01-01', fin: f.fin, zones: f.zones || [], cumul: true });
      if (colsK.length) morceaux(colsK, 18).forEach(function (part, i, all) {
        var sk = nouvelle(p.court + ' – indicateurs clés' + (all.length > 1 ? ' (' + (i + 1) + '/' + all.length + ')' : ''), coul,
          MZ ? 'Total des ' + c.ordre.length + ' zones' : '');
        tuiles(sk, part.map(function (k) {
          var t = [k.lib, f1(valK(k, c.total.valeurs), k.pct)];
          if (cC) t.push(k.stock ? 'situation du jour' : f1(valK(k, cC.total.valeurs), k.pct));
          return t;
        }), 1.5, coul);
      });
      var aires = U && R.length === 1 ? ((R[0].donnees || {}).aires || {}) : {}, nomsA = D.aires(c.ordre[0]).filter(function (a) { return aires[a]; });
      var numA = nombres.filter(function (k) { return k.ids && nomsA.some(function (a) { return k.ids.some(function (id) { return aires[a][id] != null; }); }); });
      morceaux(numA, 4).forEach(function (part) {
        if (!part.length) return;
        multiples(nouvelle(p.court + ' – indicateurs clés', coul, 'Répartition par aire de santé'), part.map(function (k) {
          return { titre: k.lib, labels: nomsA, values: nomsA.map(function (a) { var t = 0; k.ids.forEach(function (id) { t += aires[a][id] || 0; }); return t; }) };
        }), coul, { horiz: nomsA.length > 6 });
      });
      if (MZ) {
        var nzA = nombres.filter(function (k) { return c.ordre.some(function (z) { return valK(k, c.zones[z].valeurs); }); });
        morceaux(nzA.slice(0, 8), 4).forEach(function (part) {
          if (!part.length) return;
          multiples(nouvelle(p.court + ' – indicateurs clés', coul, 'Par zone de santé'), part.map(function (k) {
            return { titre: k.lib, labels: c.ordre.map(function (z) { return z.replace('Coordination provinciale', 'Prov.').replace('Boma Mangbetu', 'Boma M.'); }), values: c.ordre.map(function (z) { return z0(valK(k, c.zones[z].valeurs)); }) };
          }), coul, { horiz: c.ordre.length > 5 });
        });
      }

      // 4. Évolution sur 14 jours (courbes), à partir des rapports disponibles
      var jours = []; for (var j = 13; j >= 0; j--) jours.push(A.decaler(f.fin, -j));
      var parJour = jours.map(function (d) { return A.consolider(rapports, p, { debut: d, fin: d, zones: f.zones || [] }); });
      var i0 = parJour.findIndex(function (x) { return x.ordre.length; }); if (i0 < 0) i0 = jours.length;
      jours = jours.slice(i0); parJour = parJour.slice(i0);
      var avecDonnees = parJour.filter(function (x) { return x.ordre.length; }).length;
      var suivis = nombres.filter(function (k) { return !k.stock; }).filter(function (k) { return parJour.some(function (x) { return x.ordre.length && valK(k, x.total.valeurs); }); }).slice(0, 4);
      if (avecDonnees >= 2 && suivis.length) {
        var se = nouvelle(p.court + ' – évolution', coul, (MZ ? 'Toutes zones · du ' : 'Du ') + A.fr(jours[0]) + ' au ' + A.fr(f.fin) + ' (' + avecDonnees + ' jours avec rapport, 14 au plus)');
        multiples(se, suivis.map(function (k) { return { titre: k.lib, labels: jours.map(function (d) { return d.slice(8) + '/' + d.slice(5, 7); }), values: parJour.map(function (x) { return x.ordre.length ? z0(valK(k, x.total.valeurs)) : 0; }) }; }), coul, { line: true });
      }

      // 5. Rubriques du canevas : graphiques (listes en tableau, textes en puces)
      p.sections.forEach(function (sec) {
        if (sec.champs) {
          var ch = sec.champs.filter(function (x) { return c.ordre.some(function (z) { return c.zones[z].valeurs[x[0]] != null; }); });
          if (!ch.length) return;
          var lib = function (x) { return court(x[1], 70) + (x[2] === 'stock' ? ' *' : ''); };
          if (MZ) morceaux(ch, 4).forEach(function (part, i, all) {
            var s = nouvelle(p.court + ' – ' + sec.titre + (all.length > 1 ? ' (' + (i + 1) + '/' + all.length + ')' : ''), coul, 'Par zone de santé');
            multiples(s, part.map(function (x) { return { titre: lib(x), labels: c.ordre.map(function (z) { return z.replace('Coordination provinciale', 'Prov.').replace('Boma Mangbetu', 'Boma M.'); }), values: c.ordre.map(function (z) { return z0(c.zones[z].valeurs[x[0]]); }) }; }), coul, { horiz: c.ordre.length > 5 });
          });
          else {
            var s = nouvelle(p.court + ' – ' + sec.titre, coul, ch.some(function (x) { return x[2] === 'stock'; }) ? '* situation en fin de journée' : '');
            barres(s, [{ name: sec.titre, labels: ch.map(lib), values: ch.map(function (x) { return z0(c.total.valeurs[x[0]]); }) }], coul, { horiz: true });
          }
        } else if (sec.grille) {
          var g = sec.grille, V = c.total.valeurs;
          var lg = g.lignes.filter(function (l) { return g.colonnes.some(function (k) { return V[sec.id + '.' + l[0] + '.' + k[0]] != null; }); });
          if (!lg.length) return;
          var s2 = nouvelle(p.court + ' – ' + sec.titre, coul, [MZ ? 'Total des zones' : '', lg.some(function (l) { return l[2] === 'stock'; }) ? '* situation en fin de journée' : ''].filter(Boolean).join(' · '));
          barres(s2, g.colonnes.map(function (k) { return { name: k[1], labels: lg.map(function (l) { return court(l[1], 60) + (l[2] === 'stock' ? ' *' : ''); }), values: lg.map(function (l) { return z0(V[sec.id + '.' + l[0] + '.' + k[0]]); }) }; }), coul, { horiz: lg.length > 5 });
        } else if (sec.liste) {
          var L = A.lignesListe(R, sec.id); if (!L.length) return;
          var lc2 = sec.liste.colonnes;
          var s3 = nouvelle(p.court + ' – ' + sec.titre, coul);
          tableau(s3, (MZ ? ['Zone'] : []).concat(multi ? ['Date'] : [], lc2.map(function (k) { return k[1]; })), L.map(function (x) {
            return (MZ ? [x.zone] : []).concat(multi ? [A.fr(x.date)] : [], lc2.map(function (k) { return k[2] === 'n' ? f1(A.num(x.l[k[0]])) : (x.l[k[0]] || ''); }));
          }), { couleur: coul, wl: MZ ? 1.6 : 3, colW: !MZ && !multi && lc2.length === 2 ? [8.6, 3.73] : undefined, fs: lc2.length > 8 ? 9 : 11, gauche: lc2.every(function (k) { return k[2] !== 'n'; }) });
        } else if (sec.textes) {
          var it = [];
          sec.textes.forEach(function (t) { A.textes(R, t[0]).forEach(function (x) { it.push([(avecZone ? x.zone : '') + (sec.textes.length > 1 ? (avecZone ? ' – ' : '') + t[1] : '') + (avecZone || sec.textes.length > 1 ? ' : ' : ''), String(x.t).replace(/\s*\n\s*/g, ' ; ')]); }); });
          if (!it.length) return;
          morceaux(it, 10).forEach(function (part) { puces(nouvelle(p.court + ' – ' + sec.titre, coul), part); });
        }
      });
      // Indicateurs calculés : pourcentages en graphique, autres en chiffres
      var autres = (p.calculs || []).filter(function (k) { return !k.cle && Kt[k.id] != null; });
      if (autres.length) {
        var s4 = nouvelle(p.court + ' – indicateurs calculés', coul, MZ ? 'Total des zones' : '');
        var pc = autres.filter(function (k) { return k.pct; }), nb = autres.filter(function (k) { return !k.pct; });
        var y4 = nb.length ? tuiles(s4, nb.map(function (k) { return [k.lib, f1(Kt[k.id])]; }), 1.45, coul) + 0.1 : 1.45;
        if (pc.length) {
          if (MZ) barres(s4, pc.slice(0, 3).map(function (k) { return { name: court(k.lib, 40), labels: c.ordre.map(function (z) { return z.replace('Coordination provinciale', 'Prov.'); }), values: c.ordre.map(function (z) { var K = A.calculer(p, c.zones[z].valeurs); return K[k.id] == null ? 0 : Math.round(K[k.id] * 1000) / 10; }) }; }), coul, { y: y4, h: H - y4 - 0.5, pct: true });
          else barres(s4, [{ name: 'Taux', labels: pc.map(function (k) { return court(k.lib, 45); }), values: pc.map(function (k) { return Math.round(Kt[k.id] * 1000) / 10; }) }], coul, { y: y4, h: H - y4 - 0.5, pct: true, horiz: true });
        }
      }

      // 6. Défis et pistes de solution (toujours présentée, juste avant les activités prévues pour le lendemain)
      var Df = A.lignesListe(R, 'defis');
      var s6 = nouvelle(p.court + ' – défis et pistes de solution', coul);
      if (Df.length) tableau(s6, (MZ ? ['Zone'] : []).concat(['Défi / problème', 'Piste de solution', 'Délai']), Df.map(function (x) { return (MZ ? [x.zone] : []).concat([x.l.defi || '', x.l.solution || '', x.l.delai || '']); }),
        { couleur: coul, colW: MZ ? [1.7, 4.9, 4.9, 1.03] : [5.6, 5.6, 1.33], fs: 12, gauche: true });
      else s6.addText('Aucun défi signalé dans ' + (MZ ? 'les rapports' : 'ce rapport') + '.', { x: 0.5, y: 1.6, w: W - 1, h: 0.6, fontFace: 'Arial', fontSize: 18, color: GRIS, italic: true });

      // 7. Activités planifiées pour le jour suivant
      var pla = []; R.forEach(function (r) { A.activites(((r.donnees || {}).textes || {}).planifiees).forEach(function (a) { pla.push(avecZone ? [r.zone + ' : ', a] : ['', a]); }); });
      morceaux(pla.length ? pla : [null], 10).forEach(function (part, i, all) {
        var sp2 = nouvelle(p.court + ' – activités planifiées pour le jour suivant' + (all.length > 1 ? ' (' + (i + 1) + '/' + all.length + ')' : ''), coul, 'Programme du ' + A.fr(A.decaler(f.fin, 1)));
        puces(sp2, part[0] == null ? [] : part.map(function (a, k) { return [(i * 10 + k + 1) + '. ' + a[0], a[1]]; }), 1.5, H - 2.2);
      });
    });

    // Contrôle qualité : seulement dans le rapport de la coordination provinciale, après les activités planifiées
    var QP = root.QUALITE_PILIERS;
    if (QP && !U) {
      var RQ = QP.inspecter(rapports, f), nomQ = function (id) { return [].concat(id).map(function (x) { var p = D.pilier(x); return p ? p.court : x; }).join(' + '); };
      var sq = nouvelle('Qualité et concordance des rapports', ENCRE,
        RQ.rapports + ' rapports contrôlés · ' + RQ.sansAnomalie + ' sans anomalie · ' + RQ.erreurs + ' incohérence(s) · ' + RQ.ruptures + ' rupture(s) avec la veille · ' + RQ.ecarts + ' écart(s) entre piliers');
      var lq = RQ.concordances.filter(function (c) { return c.statut !== 'ok'; }).map(function (c) {
        return [c.zone, c.regle.lib, f1(c.a), f1(c.b), c.statut === 'ecart' ? 'Écart' : 'Pas de rapport ' + nomQ(c.manque)];
      });
      RQ.aCorriger.forEach(function (x) {
        x.internes.concat(x.continuite).forEach(function (i) { lq.push([x.r.zone, nomQ(x.r.pilier) + ' : ' + i.regle.lib, f1(i.a), f1(i.b), i.regle.grav === 'erreur' ? 'Incohérence' : 'À vérifier']); });
      });
      if (lq.length) tableau(sq, ['Zone', 'Contrôle', 'Valeur 1', 'Valeur 2', 'Statut'], lq, { colW: [1.7, 6.9, 1.0, 1.0, 1.93], fs: 11, gauche: true });
      else sq.addText('✓ Aucune incohérence ni aucun écart entre piliers sur cette période.', { x: 0.5, y: 1.6, w: W - 1, h: 0.6, fontFace: 'Arial', fontSize: 18, color: '23804F' });
    }
    // Dernière diapositive : remerciements avec l'en-tête officiel
    if (merci) { var sm = P.addSlide(); sm.background = { color: 'FFFFFF' }; sm.addImage({ data: merci, x: 0, y: 0, w: W, h: W * 708 / 1263 }); }
    var nomF = (f.pilier ? 'Rapport_' + sansAccents(D.pilier(f.pilier).court).replace(/[\s\/]+/g, '_') + '_HU_' : 'Rapports_piliers_HU_') + f.fin + (f.zone ? '_' + f.zone.replace(/\s+/g, '_') : '') + '.pptx';
    if (f.blob) return P.write({ outputType: 'blob' }).then(function (b) { return { nom: nomF, blob: b }; });
    return P.writeFile({ fileName: (f.pilier ? 'Rapport_' + sansAccents(D.pilier(f.pilier).court).replace(/[\s\/]+/g, '_') + '_HU_' : 'Rapports_piliers_HU_') + f.fin + (f.zone ? '_' + f.zone.replace(/\s+/g, '_') : '') + '.pptx' });
  }

  root.EXPORTS = { excel: excel, diapos: diapos };
})(this);
