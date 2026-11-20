/* Exports de « Rapports des piliers » : classeur Excel (SheetJS) et diapositives PowerPoint (PptxGenJS). */
(function (root) {
  'use strict';
  var D = root.PILIERS_DEF, A = root.AGREGATS;

  function filtrer(rapports, f, pilierId) {
    return A.dedoublonner(rapports).filter(function (r) {
      return (!pilierId || r.pilier === pilierId) && r.date >= f.debut && r.date <= f.fin && (!f.zones.length || f.zones.indexOf(r.zone) >= 0);
    }).sort(function (a, b) { return a.date.localeCompare(b.date) || a.zone.localeCompare(b.zone, 'fr'); });
  }
  function libPeriode(f) {
    return f.periode === '0' ? 'du ' + A.fr(f.fin) : f.periode === 'tout' ? 'cumul jusqu’au ' + A.fr(f.fin) : 'du ' + A.fr(f.debut) + ' au ' + A.fr(f.fin);
  }
  function sectionsTexte(p) { var o = []; D.sections(p).forEach(function (s) { (s.textes || []).forEach(function (t) { o.push(t); }); }); return o; }

  /* ---------- Excel ---------- */
  function excel(rapports, f) {
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
      ['SYSTÈME DE GESTION DE L’INCIDENT MVE 17'], ['17e Epidémie de la Maladie à Virus Ebola – Division Provinciale de la santé du Haut-Uele'],
      ['République Démocratique du Congo'], [], ['Rapports quotidiens des piliers'], [],
      ['Période', libPeriode(f)], ['Zone', f.zone || 'Toutes les zones'], ['Pilier', f.pilier ? D.pilier(f.pilier).nom : 'Tous'],
      ['Exporté le', new Date().toLocaleString('fr-FR')], [],
      ['Feuille « Synthèse » : consolidation par pilier et par zone. Les valeurs « fin de journée » (reports, hospitalisés, stocks…) prennent le dernier jour de la période ; les autres sont additionnées.'],
      ['Une feuille par pilier : un rapport par ligne (une ligne par zone, date et unité).'],
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
        aoa.push([A.fr(r.date), r.zone].concat(en.map(function (e) { return (d.entete || {})[e[0]] || ''; }), [r.auteur || '', r.maj ? new Date(r.maj).toLocaleString('fr-FR') : ''],
          ch.map(function (x) { return V[x.id] == null ? null : V[x.id]; }), ks.map(function (k) { return K[k.id]; }), tx.map(function (t) { return (d.textes || {})[t[0]] || ''; })));
      });
      var ws = feuille(p.court, aoa, [11, 16].concat(en.map(function () { return 16; }), [22, 18]));
      var c0 = 4 + en.length + ch.length;
      ks.forEach(function (k, j) { if (k.pct) for (var i = 1; i < aoa.length; i++) pct(ws, i, c0 + j); });
    });
    // Défis, PoE-PoC, feedbacks
    var defis = [['Pilier', 'Date', 'Zone', 'Défi / problème', 'Solution proposée', 'Responsable', 'Délai']];
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
    X.writeFile(wb, 'Rapports_piliers_HU_' + f.fin + (f.zone ? '_' + f.zone.replace(/\s+/g, '_') : '') + '.xlsx');
  }

  /* ---------- PowerPoint ---------- */
  var W = 13.333, H = 7.5, ENCRE = '0F3D3E', GRIS = '5D6B78';
  function morceaux(arr, n) { var o = []; for (var i = 0; i < arr.length; i += n) o.push(arr.slice(i, i + n)); return o.length ? o : [[]]; }
  function f1(v, pct) { if (v == null) return '–'; return pct ? (Math.round(v * 1000) / 10).toString().replace('.', ',') + ' %' : (Math.round(v * 10) / 10).toLocaleString('fr-FR'); }
  // Emblème (icons/embleme.png) en data URL pour l'intégrer à la présentation
  function chargerEmbleme(nom) {
    return fetch('icons/' + nom).then(function (r) { if (!r.ok) throw 0; return r.blob(); }).then(function (b) {
      return new Promise(function (res) { var fr = new FileReader(); fr.onload = function () { res(fr.result); }; fr.onerror = function () { res(null); }; fr.readAsDataURL(b); });
    }).catch(function () { return null; });
  }
  function diapos(rapports, f) {
    return Promise.all([chargerEmbleme('embleme.png'), chargerEmbleme('embleme-petit.png')]).then(function (e) { return diapos_(rapports, f, e[0], e[1]); });
  }
  // Couleurs de l'en-tête officiel du SGI MVE 17
  var OFF = { marine: '1F2A78', or: 'C9A227', rouge: 'E30613', vert: '00A651', bleu: '1D3F9A', rougeL: 'D52B1E', jaune: 'F2B705' };
  function diapos_(rapports, f, emb, embPetit) {
    var P = new root.PptxGenJS();
    P.layout = 'LAYOUT_WIDE'; P.author = 'DPS Haut-Uélé – SGI MVE'; P.title = 'Rapports des piliers ' + libPeriode(f);
    var periode = libPeriode(f), lieu = f.zone || 'Province du Haut-Uélé';
    function nouvelle(titre, couleur, sousTitre) {
      var s = P.addSlide();
      s.background = { color: 'FFFFFF' };
      s.addShape(P.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.85, fill: { color: couleur || ENCRE }, line: { color: couleur || ENCRE } });
      s.addText(titre, { x: 0.4, y: 0.1, w: W - 3.2, h: 0.65, fontFace: 'Arial', fontSize: 22, bold: true, color: 'FFFFFF', valign: 'middle', fit: 'shrink' });
      s.addText(lieu + '\n' + periode, { x: W - 2.9, y: 0.1, w: 2.6, h: 0.65, fontFace: 'Arial', fontSize: 10, color: 'FFFFFF', align: 'right', valign: 'middle' });
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

    // Couverture
    // Couverture : en-tête officiel (Task Force, Ministère, SGI MVE 17, DPS Haut-Uele)
    var c0 = P.addSlide();
    c0.background = { color: 'FFFFFF' };
    var losange = function (x, y, t, coul, transp) {
      c0.addShape(P.ShapeType.rect, { x: x, y: y, w: t, h: t, rotate: 45, fill: { color: coul, transparency: transp }, line: { color: coul, transparency: 100 } });
    };
    losange(-0.55, -0.75, 1.9, '7FA3BA', 55); losange(0.6, 0.05, 1.0, '5E8AA6', 45);
    losange(W - 1.45, -0.8, 2.1, '8FD0EC', 50); losange(W - 2.45, 0.5, 1.1, '5BBDE6', 45);
    if (emb) c0.addImage({ data: emb, x: W / 2 - 0.82, y: 0.12, w: 1.64, h: 1.55 });
    var cf = 'Century Gothic';
    c0.addText('TASK FORCE PRÉSIDENTIELLE EBOLA 17', { x: 0.5, y: 1.72, w: W - 1, h: 0.38, fontFace: cf, fontSize: 17, color: OFF.marine, align: 'center', charSpacing: 1.5 });
    c0.addText('MINISTÈRE DE LA SANTÉ PUBLIQUE, HYGIÈNE ET PRÉVOYANCE SOCIALE', { x: 0.5, y: 2.08, w: W - 1, h: 0.34, fontFace: cf, fontSize: 13, color: OFF.marine, align: 'center', charSpacing: 1 });
    c0.addText('SYSTÈME DE GESTION DE L’INCIDENT MVE 17', { x: 0.5, y: 2.42, w: W - 1, h: 0.45, fontFace: cf, fontSize: 22, color: OFF.or, align: 'center', charSpacing: 2 });
    [OFF.bleu, OFF.rougeL, OFF.jaune].forEach(function (c, i) { c0.addShape(P.ShapeType.rect, { x: 3.35 + i * 2.25, y: 2.93, w: 2.15, h: 0.05, fill: { color: c }, line: { color: c } }); });
    c0.addText([{ text: '17', options: {} }, { text: 'e', options: { superscript: true } }, { text: ' Epidémie de la Maladie à Virus Ebola', options: {} }],
      { x: 0.5, y: 3.15, w: W - 1, h: 0.6, fontFace: cf, fontSize: 30, color: OFF.rouge, align: 'center' });
    c0.addText('Division Provinciale de la santé du Haut-Uele', { x: 0.5, y: 3.72, w: W - 1, h: 0.5, fontFace: cf, fontSize: 24, color: OFF.vert, align: 'center' });
    c0.addText('République Démocratique du Congo', { x: 0.5, y: 4.2, w: W - 1, h: 0.35, fontFace: cf, fontSize: 14, bold: true, underline: { style: 'sng' }, color: '000000', align: 'center' });
    c0.addShape(P.ShapeType.rect, { x: 2.4, y: 4.95, w: W - 4.8, h: 1.25, fill: { color: 'F1F5F6' }, line: { color: 'F1F5F6' } });
    c0.addText('Rapports quotidiens des piliers – réunion de coordination', { x: 2.4, y: 5.02, w: W - 4.8, h: 0.55, fontFace: cf, fontSize: 20, bold: true, color: ENCRE, align: 'center' });
    c0.addText(lieu + ' · ' + periode, { x: 2.4, y: 5.55, w: W - 4.8, h: 0.5, fontFace: cf, fontSize: 16, color: OFF.or, bold: true, align: 'center' });
    c0.addText('Généré le ' + new Date().toLocaleString('fr-FR') + ' à partir des rapports saisis dans l’application « Rapports des piliers ».', { x: 0.6, y: H - 0.6, w: W - 1.2, h: 0.35, fontFace: 'Arial', fontSize: 10, color: '8A97A3', align: 'center' });

    // Complétude
    var C = A.completude(rapports, f.fin), cols = D.ZONES.concat([D.PROVINCE]);
    var sc = nouvelle('Complétude des rapports du ' + A.fr(f.fin));
    tableau(sc, ['Pilier'].concat(cols.map(function (z) { return z === D.PROVINCE ? 'Prov.' : z.replace('Boma Mangbetu', 'Boma M.'); }), ['ZS']),
      D.PILIERS.map(function (p) { var c = C[p.id] || {}; return [p.court].concat(cols.map(function (z) { return c[z] ? '✓' : ''; }), [D.ZONES.filter(function (z) { return c[z]; }).length + '/13']); }),
      { wl: 2.4, fs: 10, gauche: false });

    // Qualité et concordance
    var QP = root.QUALITE_PILIERS;
    if (QP) {
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

    D.PILIERS.filter(function (p) { return !f.pilier || p.id === f.pilier; }).forEach(function (p) {
      var c = A.consolider(rapports, p, f); if (!c.ordre.length) return;
      var coul = p.couleur.replace('#', ''), R = c.total.rapports, Kt = A.calculer(p, c.total.valeurs);
      // Points saillants + indicateurs clés
      var s1 = nouvelle(p.nom, coul, 'Points saillants · ' + c.ordre.length + ' zone' + (c.ordre.length > 1 ? 's' : '') + ' ayant rapporté');
      var cles = A.colonnesCles(p).map(function (col) { return [col.lib, f1(A.valCol(col, c.total.valeurs, Kt), col.pct)]; }).slice(0, 6);
      var bw = (W - 0.8 - 0.15 * (cles.length - 1)) / Math.max(1, cles.length);
      cles.forEach(function (k, i) {
        var x = 0.4 + i * (bw + 0.15);
        s1.addShape(P.ShapeType.roundRect, { x: x, y: 1.45, w: bw, h: 1.15, fill: { color: 'F1F4F5' }, line: { color: 'F1F4F5' }, rectRadius: 0.08 });
        s1.addText(k[1], { x: x, y: 1.5, w: bw, h: 0.6, fontFace: 'Arial', fontSize: 24, bold: true, color: coul, align: 'center' });
        s1.addText(k[0], { x: x + 0.05, y: 2.05, w: bw - 0.1, h: 0.5, fontFace: 'Arial', fontSize: 10, color: GRIS, align: 'center', valign: 'top', fit: 'shrink' });
      });
      puces(s1, textesEnPuces(A.textes(R, 'saillants'), multi), 2.85, H - 3.4);
      // Indicateurs clés par zone
      var colsK = A.colonnesCles(p);
      var lig = function (nom, V) { var K = A.calculer(p, V); return [nom].concat(colsK.map(function (col) { return f1(A.valCol(col, V, K), col.pct); })); };
      if (c.ordre.length > 1 || colsK.length > 6) morceaux(colsK, 8).forEach(function (part, i, all) {
        var save = colsK; colsK = part;
        var s = nouvelle(p.nom + ' – indicateurs clés par zone' + (all.length > 1 ? ' (' + (i + 1) + '/' + all.length + ')' : ''), coul);
        tableau(s, ['Zone'].concat(part.map(function (x) { return x.lib; })), c.ordre.map(function (z) { return lig(z, c.zones[z].valeurs); }).concat(c.ordre.length > 1 ? [lig('Total', c.total.valeurs)] : []), { couleur: coul, wl: 2.2 });
        colsK = save;
      });
      // Sections du canevas
      p.sections.forEach(function (sec) {
        if (sec.champs) {
          var ch = sec.champs.filter(function (x) { return c.ordre.some(function (z) { return c.zones[z].valeurs[x[0]] != null; }); });
          if (!ch.length) return;
          morceaux(ch, 7).forEach(function (part, i, all) {
            var s = nouvelle(p.court + ' – ' + sec.titre + (all.length > 1 ? ' (' + (i + 1) + '/' + all.length + ')' : ''), coul);
            var rows = c.ordre.map(function (z) { return [z].concat(part.map(function (x) { return f1(c.zones[z].valeurs[x[0]]); })); });
            if (c.ordre.length > 1) rows.push(['Total'].concat(part.map(function (x) { return f1(c.total.valeurs[x[0]]); })));
            tableau(s, ['Zone'].concat(part.map(function (x) { return x[1] + (x[2] === 'stock' ? ' *' : ''); })), rows, { couleur: coul, wl: 2.2 });
          });
        } else if (sec.grille) {
          var g = sec.grille, V = c.total.valeurs;
          var any = g.lignes.some(function (l) { return g.colonnes.some(function (k) { return V[sec.id + '.' + l[0] + '.' + k[0]] != null; }); });
          if (!any) return;
          var s = nouvelle(p.court + ' – ' + sec.titre, coul, c.ordre.length > 1 ? 'Total des zones : ' + c.ordre.join(', ') : c.ordre[0]);
          tableau(s, [''].concat(g.colonnes.map(function (k) { return k[1]; })), g.lignes.map(function (l) {
            return [l[1] + (l[2] === 'stock' ? ' *' : '')].concat(g.colonnes.map(function (k) { return f1(V[sec.id + '.' + l[0] + '.' + k[0]]); }));
          }), { couleur: coul, y: 1.5, wl: 4 });
        } else if (sec.liste) {
          var L = A.lignesListe(R, sec.id); if (!L.length) return;
          var lc2 = sec.liste.colonnes;
          var s3 = nouvelle(p.court + ' – ' + sec.titre, coul);
          tableau(s3, ['Zone'].concat(multi ? ['Date'] : [], lc2.map(function (k) { return k[1]; })), L.map(function (x) {
            return [x.zone].concat(multi ? [A.fr(x.date)] : [], lc2.map(function (k) { return k[2] === 'n' ? f1(A.num(x.l[k[0]])) : (x.l[k[0]] || ''); }));
          }), { couleur: coul, wl: 1.6, fs: lc2.length > 8 ? 9 : 11, gauche: lc2.every(function (k) { return k[2] !== 'n'; }) });
        } else if (sec.textes) {
          var it = [];
          sec.textes.forEach(function (t) { A.textes(R, t[0]).forEach(function (x) { it.push([x.zone + (sec.textes.length > 1 ? ' – ' + t[1] : '') + ' : ', String(x.t).replace(/\s*\n\s*/g, ' ; ')]); }); });
          if (!it.length) return;
          morceaux(it, 10).forEach(function (part) { puces(nouvelle(p.court + ' – ' + sec.titre, coul), part); });
        }
      });
      // Indicateurs calculés non clés
      var autres = (p.calculs || []).filter(function (k) { return !k.cle; });
      if (autres.length && autres.some(function (k) { return Kt[k.id] != null; })) {
        var s4 = nouvelle(p.court + ' – indicateurs calculés', coul);
        tableau(s4, ['Zone'].concat(autres.map(function (k) { return k.lib; })), c.ordre.map(function (z) { var K = A.calculer(p, c.zones[z].valeurs); return [z].concat(autres.map(function (k) { return f1(K[k.id], k.pct); })); })
          .concat(c.ordre.length > 1 ? [['Total'].concat(autres.map(function (k) { return f1(Kt[k.id], k.pct); }))] : []), { couleur: coul, wl: 2.2 });
      }
      // Activités, défis
      var re = textesEnPuces(A.textes(R, 'realisees'), multi), pl = textesEnPuces(A.textes(R, 'planifiees'), multi);
      if (re.length || pl.length) {
        var s5 = nouvelle(p.court + ' – activités', coul);
        s5.addText('Réalisées', { x: 0.5, y: 1.0, w: 6, h: 0.4, fontFace: 'Arial', fontSize: 15, bold: true, color: coul });
        s5.addText('Planifiées', { x: 6.9, y: 1.0, w: 6, h: 0.4, fontFace: 'Arial', fontSize: 15, bold: true, color: coul });
        puces(s5, re, 1.45, H - 2.1, 0.5, 6.1); puces(s5, pl, 1.45, H - 2.1, 6.9, 6);
      }
      var Df = A.lignesListe(R, 'defis');
      if (Df.length) {
        var s6 = nouvelle(p.court + ' – défis et actions correctrices', coul);
        tableau(s6, ['Zone', 'Défi / problème', 'Solution proposée', 'Responsable', 'Délai'], Df.map(function (x) { return [x.zone, x.l.defi || '', x.l.solution || '', x.l.resp || '', x.l.delai || '']; }),
          { couleur: coul, colW: [1.6, 4.2, 4.2, 1.6, 0.93], fs: 11, gauche: true });
      }
    });
    return P.writeFile({ fileName: 'Rapports_piliers_HU_' + f.fin + (f.zone ? '_' + f.zone.replace(/\s+/g, '_') : '') + '.pptx' });
  }

  root.EXPORTS = { excel: excel, diapos: diapos };
})(this);
