/* Application « Rapports des piliers MVE » : saisie hors ligne, synchronisation Google Sheets, tableau de bord. */
(function () {
  'use strict';
  var D = window.PILIERS_DEF, A = window.AGREGATS, Q = window.QUALITE_PILIERS;
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var $$ = function (s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); };
  var nf = new Intl.NumberFormat('fr-FR', { maximumFractionDigits: 1 });
  var pf = new Intl.NumberFormat('fr-FR', { style: 'percent', maximumFractionDigits: 1 });
  function fmt(v, pct) { return v == null ? '–' : pct ? pf.format(v) : nf.format(v); }
  function el(tag, props, kids) {
    var e = document.createElement(tag);
    if (props) Object.keys(props).forEach(function (k) {
      if (k === 'text') e.textContent = props[k]; else if (k === 'cls') e.className = props[k];
      else if (k.slice(0, 2) === 'on') e.addEventListener(k.slice(2), props[k]); else if (props[k] != null) e.setAttribute(k, props[k]);
    });
    (kids || []).forEach(function (c) { if (c != null) e.append(c); });
    return e;
  }

  /* ---------- Stockage local (IndexedDB) ---------- */
  var dbP = null;
  function db() {
    return dbP || (dbP = new Promise(function (res, rej) {
      var rq = indexedDB.open('rapports-piliers', 1);
      rq.onupgradeneeded = function () { rq.result.createObjectStore('rapports', { keyPath: 'id' }); rq.result.createObjectStore('kv'); };
      rq.onsuccess = function () { res(rq.result); }; rq.onerror = function () { rej(rq.error); };
    }));
  }
  function tx(store, mode, fn) {
    return db().then(function (d) {
      return new Promise(function (res, rej) {
        var t = d.transaction(store, mode), out = fn(t.objectStore(store));
        t.oncomplete = function () { res(out && out.result !== undefined ? out.result : undefined); }; t.onerror = function () { rej(t.error); };
      });
    });
  }
  var S = {
    tous: function () { return tx('rapports', 'readonly', function (s) { return s.getAll(); }); },
    put: function (r) { return tx('rapports', 'readwrite', function (s) { s.put(r); }); },
    putMany: function (L) { return tx('rapports', 'readwrite', function (s) { L.forEach(function (r) { s.put(r); }); }); },
    get: function (k) { return tx('kv', 'readonly', function (s) { return s.get(k); }); },
    set: function (k, v) { return tx('kv', 'readwrite', function (s) { s.put(v, k); }); }
  };
  function lsGet(k) { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } }
  function lsSet(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* indisponible */ } }

  var state = { rapports: [], reglages: {}, vue: 'saisir', courant: null, sale: false, listes: {}, sync: false };
  var toastT;
  function toast(m, ms) { var t = $('#toast'); t.textContent = m; t.hidden = false; clearTimeout(toastT); toastT = setTimeout(function () { t.hidden = true; }, ms || 3500); }
  function busy(on, txt) { $('#busy').hidden = !on; if (txt) $('#busy-txt').textContent = txt; }
  function uid() { return 'r' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8); }
  function maintenant() { return new Date().toISOString(); }

  /* ---------- Listes déroulantes ---------- */
  function remplir(sel, items, val, vide) {
    sel.replaceChildren();
    if (vide) sel.appendChild(new Option(vide, ''));
    items.forEach(function (it) { var o = Array.isArray(it) ? new Option(it[1], it[0]) : new Option(it, it); sel.appendChild(o); });
    if (val != null) sel.value = val;
  }
  var toutesZones = D.ZONES.concat([D.PROVINCE, D.HORS]);
  var listePiliers = D.PILIERS.map(function (p) { return [p.id, p.nom]; });

  /* ---------- Formulaire ---------- */
  function pilierCourant() { return D.pilier($('#f-pilier').value) || D.PILIERS[0]; }
  function enteteCourante(p) {
    var e = {}; (p.entete || []).forEach(function (c) { e[c[0]] = ($('#e-' + c[0]) || {}).value || ''; }); return e;
  }
  function trouverRapport(p, date, zone, entete) {
    var cle = A.cleRapport({ pilier: p.id, date: date, zone: zone, donnees: { entete: entete } });
    var c = state.rapports.filter(function (r) { return A.cleRapport(r) === cle && !(r.donnees && r.donnees.supprime); })
      .sort(function (a, b) { return String(b.maj).localeCompare(String(a.maj)); });
    return c[0] || null;
  }
  function construireEntete(p, entete) {
    var host = $('#f-entete'); host.replaceChildren();
    (p.entete || []).forEach(function (c) {
      var f = el('div', { cls: 'field' }), id = 'e-' + c[0];
      f.append(el('label', { for: id, text: c[1] }));
      var inp;
      if (c[2] === 'choix') { inp = el('select', { id: id }); remplir(inp, c[3], null, 'Choisir…'); }
      else inp = el('input', { id: id, placeholder: c[0] === 'unite' ? 'ex. : CTE HGR Isiro' : '' });
      inp.value = (entete && entete[c[0]]) || '';
      inp.addEventListener('change', function () { chargerSelection(); });
      f.append(inp); host.append(f);
    });
    host.hidden = !(p.entete || []).length;
  }
  function inputNum(id, val) {
    var i = el('input', { cls: 'num', inputmode: 'decimal', autocomplete: 'off', 'data-id': id, 'aria-label': id });
    i.value = val == null ? '' : String(val).replace('.', ',');
    i.addEventListener('input', function () { verifNum(i); state.sale = true; majCalculs(); });
    return i;
  }
  function verifNum(i) { var v = i.value.trim(); i.classList.toggle('err', v !== '' && (A.num(v) == null || A.num(v) < 0)); }

  function construireFormulaire(p, rap) {
    var d = (rap && rap.donnees) || {}, V = d.valeurs || {}, T = d.textes || {};
    state.listes = JSON.parse(JSON.stringify(d.listes || {}));
    var host = $('#f-sections'); host.replaceChildren();
    D.sections(p).forEach(function (sec, i) {
      var det = el('details', { cls: 'sec' }); if (i < 2) det.open = true;
      var sum = el('summary', {}, [document.createTextNode(sec.titre), el('span', { cls: 'cnt', 'data-cnt': sec.id })]);
      var body = el('div', { cls: 'sec-body' });
      det.append(sum, body);
      if (sec.champs) {
        var g = el('div', { cls: 'nums' });
        sec.champs.forEach(function (c) {
          var inp = inputNum(c[0], V[c[0]]); inp.id = 'n-' + c[0];
          g.append(el('div', { cls: 'nrow' }, [el('label', { for: inp.id, text: c[1] + (c[2] === 'stock' ? ' *' : '') }), inp]));
        });
        body.append(g);
        if (sec.champs.some(function (c) { return c[2] === 'stock'; })) body.append(el('p', { cls: 'note', text: '* Situation en fin de journée (non cumulée d’un jour à l’autre).' }));
      }
      if (sec.grille) {
        var tb = el('table', { cls: 'g' }), hd = el('tr', {}, [el('th', { cls: 'l', text: '' })]);
        sec.grille.colonnes.forEach(function (c) { hd.append(el('th', { text: c[1] })); });
        tb.append(el('thead', {}, [hd]));
        var bd = el('tbody');
        sec.grille.lignes.forEach(function (l) {
          var tr = el('tr', {}, [el('td', { cls: 'l', text: l[1] + (l[2] === 'stock' ? ' *' : '') })]);
          sec.grille.colonnes.forEach(function (c) { var id = sec.id + '.' + l[0] + '.' + c[0]; var inp = inputNum(id, V[id]); inp.setAttribute('aria-label', l[1] + ' – ' + c[1]); tr.append(el('td', {}, [inp])); });
          bd.append(tr);
        });
        tb.append(bd); body.append(el('div', { cls: 'gwrap' }, [tb]));
        if (sec.grille.lignes.some(function (l) { return l[2] === 'stock'; })) body.append(el('p', { cls: 'note', text: '* Situation en fin de journée (non cumulée d’un jour à l’autre).' }));
      }
      if (sec.textes) sec.textes.forEach(function (t) {
        var ta = el('textarea', { id: 't-' + t[0], 'data-txt': t[0] }); ta.value = T[t[0]] || '';
        ta.addEventListener('input', function () { state.sale = true; });
        body.append(el('div', { cls: 'field' }, [el('label', { for: ta.id, text: t[1] }), ta]));
      });
      if (sec.liste) {
        var wrap = el('div', { cls: 'stack', 'data-liste': sec.id });
        body.append(wrap, el('button', { cls: 'btn ghost small', type: 'button', text: '+ Ajouter une ligne', onclick: function () {
          (state.listes[sec.id] = state.listes[sec.id] || []).push({}); state.sale = true; dessinerListe(sec); } }));
        if (sec.id === 'pts' && !state.listes.pts) {
          var pts = D.POINTS[$('#f-zone').value] || [];
          if (pts.length) state.listes.pts = pts.map(function (n) { return { nom: n }; });
        }
        dessinerListe(sec);
      }
      host.append(det);
    });
    majCalculs();
  }
  function dessinerListe(sec) {
    var wrap = $('[data-liste="' + sec.id + '"]'); if (!wrap) return;
    wrap.replaceChildren();
    var L = state.listes[sec.id] || [];
    if (!L.length) wrap.append(el('p', { cls: 'note', text: 'Aucune ligne.' }));
    L.forEach(function (row, i) {
      var box = el('div', { cls: 'liste-row' });
      box.append(el('button', { cls: 'del', type: 'button', 'aria-label': 'Supprimer la ligne', text: '×', onclick: function () { L.splice(i, 1); state.sale = true; dessinerListe(sec); majCalculs(); } }));
      sec.liste.colonnes.forEach(function (c) {
        var f = el('div', { cls: 'field' + (c[2] === 'texte' && c[0] !== 'nom' ? ' wide' : '') }), inp;
        if (c[2] === 'choix') { inp = el('select'); remplir(inp, c[3], null, '—'); }
        else if (c[2] === 'n') { inp = el('input', { inputmode: 'decimal', autocomplete: 'off' }); inp.style.textAlign = 'right'; }
        else if (c[0] === 'nom') { inp = el('input', { list: 'dl-points' }); }
        else inp = el('textarea', { rows: '2' });
        inp.value = row[c[0]] == null ? '' : row[c[0]];
        inp.addEventListener('input', function () { row[c[0]] = inp.value; state.sale = true; majCalculs(); });
        inp.addEventListener('change', function () { row[c[0]] = inp.value; majCalculs(); });
        f.append(el('label', { text: c[1] }), inp); box.append(f);
      });
      wrap.append(box);
    });
    var dl = $('#dl-points'); if (!dl) { dl = el('datalist', { id: 'dl-points' }); document.body.append(dl); }
    dl.replaceChildren.apply(dl, (D.POINTS[$('#f-zone').value] || []).map(function (n) { return new Option(n); }));
  }
  function lireFormulaire(p) {
    var valeurs = {}, textes = {}, listes = {};
    $$('#f-sections input.num').forEach(function (i) { var v = A.num(i.value.trim()); if (i.value.trim() !== '' && v != null) valeurs[i.dataset.id] = v; });
    $$('#f-sections textarea[data-txt]').forEach(function (t) { if (t.value.trim()) textes[t.dataset.txt] = t.value.trim(); });
    Object.keys(state.listes).forEach(function (k) {
      var L = (state.listes[k] || []).filter(function (x) { return Object.keys(x).some(function (c) { return x[c] != null && String(x[c]).trim() !== ''; }); });
      if (L.length) listes[k] = L;
    });
    return { valeurs: valeurs, textes: textes, listes: listes, entete: enteteCourante(p) };
  }
  function majCalculs() {
    var p = pilierCourant(), d = lireFormulaire(p), tmp = { donnees: d }, V = {};
    A.ids(p).forEach(function (c) { var v = A.valeur(tmp, c.id); if (v != null) V[c.id] = v; });
    var K = A.calculer(p, V), host = $('#f-calc');
    host.replaceChildren.apply(host, (p.calculs || []).map(function (c) {
      return el('div', {}, [el('span', { text: c.lib }), el('b', { text: fmt(K[c.id], c.pct) })]);
    }));
    $('#f-calc-card').hidden = !(p.calculs || []).length;
    majControles(p, d);
    // Nombre de valeurs saisies par section
    D.sections(p).forEach(function (sec) {
      var c = $('[data-cnt="' + sec.id + '"]'); if (!c) return;
      var n = $$('details.sec input.num', c.closest('details')).filter(function (i) { return i.value.trim() !== ''; }).length
        + $$('textarea[data-txt]', c.closest('details')).filter(function (t) { return t.value.trim(); }).length
        + (sec.liste ? (d.listes[sec.id] || []).length : 0);
      c.textContent = n ? '· ' + n + ' renseigné' + (n > 1 ? 's' : '') : '';
    });
  }
  function cleSelection() {
    var p = pilierCourant(); return A.cleRapport({ pilier: p.id, date: $('#f-date').value, zone: $('#f-zone').value, donnees: { entete: enteteCourante(p) } });
  }
  function chargerSelection(force) {
    if (!force && cleSelection() === state._cleAff) return; // rien n'a changé (ex. : champ quitté sans modification)
    if (!force && state.sale && !confirm('Des valeurs saisies ne sont pas enregistrées. Les abandonner ?')) {
      restaurerSelection(); return;
    }
    var p = pilierCourant();
    if (state._pilierAff !== p.id) { construireEntete(p, state.courant && state.courant.pilier === p.id ? state.courant.donnees.entete : null); state._pilierAff = p.id; }
    var zones = D.zonesPour(p), zv = $('#f-zone').value;
    remplir($('#f-zone'), zones, zones.indexOf(zv) >= 0 ? zv : (zones.indexOf(state.reglages.zone) >= 0 ? state.reglages.zone : zones[0]));
    $('#l-zone').textContent = p.zoneLib || 'Zone de santé';
    var rap = trouverRapport(p, $('#f-date').value, $('#f-zone').value, enteteCourante(p));
    state.courant = rap;
    construireFormulaire(p, rap);
    state.sale = false; memoSelection(); state._cleAff = cleSelection();
    var e = $('#f-etat');
    if (rap) e.textContent = 'Rapport existant' + (rap.auteur ? ' de ' + rap.auteur : '') + ' · ' + libEtat(rap) + '. Vos modifications le remplaceront.';
    else e.textContent = 'Nouveau rapport.';
    $('#btn-supprimer').hidden = !rap;
  }
  var selMemo = {};
  function memoSelection() { selMemo = { p: $('#f-pilier').value, d: $('#f-date').value, z: $('#f-zone').value }; }
  function restaurerSelection() { $('#f-pilier').value = selMemo.p; $('#f-date').value = selMemo.d; $('#f-zone').value = selMemo.z; }
  function libEtat(r) { return r._etat === 'brouillon' ? 'brouillon (non envoyé)' : r._etat === 'a_envoyer' ? 'en attente d’envoi' : 'envoyé'; }

  /* Contrôles en direct pendant la saisie (cohérence interne + continuité avec la veille) */
  function rapportFormulaire(p, d) { return { pilier: p.id, date: $('#f-date').value, zone: $('#f-zone').value, donnees: d }; }
  function majControles(p, d) {
    var r = rapportFormulaire(p, d), I = Q.controlerRapport(r), T = Q.controlerContinuite(r, state.rapports), host = $('#f-ctrl');
    var ko = I.filter(function (x) { return !x.ok; }), kt = T.filter(function (x) { return !x.ok; });
    host.replaceChildren();
    ko.forEach(function (x) { host.append(el('li', { cls: x.regle.grav }, [el('b', { text: (x.regle.grav === 'erreur' ? 'Incohérence : ' : 'À vérifier : ') + x.regle.lib }), x.a != null || x.b != null ? el('small', { text: 'Valeurs : ' + Q.nb(x.a) + ' / ' + Q.nb(x.b) }) : null])); });
    kt.forEach(function (x) { host.append(el('li', { cls: 'alerte' }, [el('b', { text: 'Continuité : ' + x.regle.lib }), el('small', { text: 'Saisi : ' + Q.nb(x.a) + ' · d’après le rapport d’hier : ' + Q.nb(x.b) })])); });
    var n = I.length + T.length;
    if (!ko.length && !kt.length) host.append(el('li', { cls: 'ok', text: n ? '✓ ' + n + ' contrôle' + (n > 1 ? 's' : '') + ' respecté' + (n > 1 ? 's' : '') + '.' : 'Les contrôles s’afficheront au fil de la saisie.' }));
    // Reports proposés d'après la veille
    var RP = Q.reportsProposes(r, state.rapports), box = $('#f-reports');
    var diff = RP && Object.keys(RP.valeurs).filter(function (id) { return d.valeurs[id] !== RP.valeurs[id]; });
    box.hidden = !(diff && diff.length);
    if (diff && diff.length) {
      var ch = D.champs(p);
      box.replaceChildren(el('span', { text: 'D’après le rapport d’hier (' + A.fr(RP.veille.date) + ') : ' + diff.map(function (id) {
        var c = ch.find(function (x) { return x.id === id; }); return (c ? c.lib : id) + ' = ' + Q.nb(RP.valeurs[id]); }).join(' ; ') + '.' }),
        el('button', { cls: 'btn ghost small', type: 'button', text: 'Reprendre ces reports', onclick: function () {
          diff.forEach(function (id) { var i = $('input.num[data-id="' + id + '"]'); if (i) { i.value = String(RP.valeurs[id]).replace('.', ','); var dt = i.closest('details'); if (dt) dt.open = true; } });
          state.sale = true; majCalculs(); toast('Reports repris du rapport d’hier.'); } }));
    }
    state.ctrlErreurs = ko.filter(function (x) { return x.regle.grav === 'erreur'; }).length;
  }
  function enregistrer(envoyer) {
    var p = pilierCourant(), d = lireFormulaire(p);
    if (envoyer && state.ctrlErreurs && !confirm(state.ctrlErreurs + ' incohérence' + (state.ctrlErreurs > 1 ? 's' : '') + ' dans ce rapport (voir « Contrôles de cohérence »). Envoyer quand même ?')) {
      $('#f-ctrl-card').scrollIntoView({ behavior: 'smooth', block: 'center' }); return;
    }
    if ($$('#f-sections input.num.err').length) { toast('Corrigez les valeurs en rouge (nombres positifs uniquement).'); $('#f-sections input.num.err').focus(); return; }
    if ((p.entete || []).some(function (c) { return !d.entete[c[0]]; })) { toast('Indiquez : ' + p.entete.map(function (c) { return c[1]; }).join(', ') + '.'); return; }
    if (!Object.keys(d.valeurs).length && !Object.keys(d.textes).length && !Object.keys(d.listes).length) { toast('Le rapport est vide.'); return; }
    var ex = state.courant, r = {
      id: ex ? ex.id : uid(), pilier: p.id, date: $('#f-date').value, zone: $('#f-zone').value,
      auteur: ex && ex.auteur ? ex.auteur : (state.reglages.auteur || ''), maj: maintenant(), donnees: d,
      _etat: envoyer ? 'a_envoyer' : 'brouillon', _local: true
    };
    if (ex && ex.auteur && state.reglages.auteur && ex.auteur !== state.reglages.auteur) r.donnees.modifiePar = state.reglages.auteur;
    S.put(r).then(function () {
      remplacer(r); state.courant = r; state.sale = false;
      $('#f-etat').textContent = 'Rapport ' + libEtat(r) + '.'; $('#btn-supprimer').hidden = false;
      majBadge();
      if (envoyer) { toast(navigator.onLine ? 'Rapport validé : envoi en cours…' : 'Rapport validé : il partira au retour du réseau.'); synchroniser(true); }
      else toast('Rapport enregistré sur le téléphone (pas encore envoyé).');
    });
  }
  function supprimer() {
    var r = state.courant; if (!r || !confirm('Supprimer ce rapport ? Il sera aussi retiré de Google Sheets à la prochaine synchronisation.')) return;
    var s = Object.assign({}, r, { maj: maintenant(), donnees: Object.assign({}, r.donnees, { supprime: true }), _etat: r._etat === 'brouillon' && !r._envoyeUneFois ? 'brouillon' : 'a_envoyer' });
    S.put(s).then(function () { remplacer(s); toast('Rapport supprimé.'); chargerSelection(true); majBadge(); synchroniser(true); });
  }
  function remplacer(r) {
    var i = state.rapports.findIndex(function (x) { return x.id === r.id; });
    if (i >= 0) state.rapports[i] = r; else state.rapports.push(r);
  }

  /* ---------- Synchronisation Google Sheets ---------- */
  function appel(corps) {
    var R = state.reglages;
    return fetch(R.url, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(Object.assign({ cle: R.cle }, corps)), redirect: 'follow' })
      .then(function (res) { if (!res.ok) throw new Error('Le script Google a répondu ' + res.status + '.'); return res.json(); })
      .then(function (j) {
        if (!j || !j.ok) throw new Error(j && j.erreur === 'cle' ? 'Clé d’équipe refusée.' : 'Erreur du script Google : ' + ((j && j.erreur) || 'réponse inattendue') + '.');
        return j;
      });
  }
  function nettoyer(r) { return { id: r.id, pilier: r.pilier, date: r.date, zone: r.zone, auteur: r.auteur, maj: r.maj, donnees: r.donnees }; }
  function synchroniser(silencieux, tout) {
    var R = state.reglages;
    if (!R.url || !R.cle) { if (!silencieux) { toast('Indiquez d’abord l’adresse du script et la clé d’équipe.'); ouvrirReglages(); } majEtat(); return Promise.resolve(); }
    if (state.sync) return Promise.resolve();
    state.sync = true; $('#btn-sync').classList.add('spin'); majEtat();
    var attente = state.rapports.filter(function (r) { return r._etat === 'a_envoyer'; }), envoyes = 0, recus = 0;
    var lots = []; for (var i = 0; i < attente.length; i += 25) lots.push(attente.slice(i, i + 25));
    var chaine = lots.reduce(function (pr, lot) {
      return pr.then(function () {
        return appel({ action: 'envoyer', rapports: lot.map(nettoyer) }).then(function (j) {
          var ok = {}; (j.ids || []).forEach(function (id) { ok[id] = 1; });
          var maj = lot.filter(function (r) { return ok[r.id]; }).map(function (r) { return Object.assign({}, r, { _etat: 'envoye', _envoyeUneFois: true }); });
          envoyes += maj.length; maj.forEach(remplacer); return S.putMany(maj);
        });
      });
    }, Promise.resolve());
    return chaine.then(function () { return tout ? '' : S.get('depuis'); }).then(function (depuis) {
      return appel({ action: 'lire', depuis: depuis || '' });
    }).then(function (j) {
      var nouv = [];
      (j.rapports || []).forEach(function (r) {
        var ex = state.rapports.find(function (x) { return x.id === r.id; });
        if (ex && (ex._etat !== 'envoye') && String(ex.maj) > String(r.maj)) return; // modification locale plus récente
        if (!ex || String(r.maj) >= String(ex.maj)) { var n = Object.assign({}, r, { _etat: 'envoye', _local: ex ? ex._local : false, _envoyeUneFois: true }); nouv.push(n); remplacer(n); }
      });
      recus = nouv.length;
      return S.putMany(nouv).then(function () { return S.set('depuis', j.maintenant || ''); }).then(function () { return S.set('derniereSync', Date.now()); });
    }).then(function () {
      state.derniere = Date.now();
      if (!silencieux || envoyes) toast((envoyes ? envoyes + ' rapport' + (envoyes > 1 ? 's' : '') + ' envoyé' + (envoyes > 1 ? 's' : '') + '. ' : '') +
        (recus ? recus + ' rapport' + (recus > 1 ? 's' : '') + ' reçu' + (recus > 1 ? 's' : '') + '.' : (envoyes ? '' : 'Tout est à jour.')));
      if (recus || envoyes) rafraichir();
    }).catch(function (e) {
      state.erreur = /Failed to fetch|NetworkError|Load failed/i.test(e.message) ? 'Pas de réseau : les rapports partiront plus tard.' : e.message;
      if (!silencieux) toast(state.erreur, 6000);
    }).then(function () { state.sync = false; $('#btn-sync').classList.remove('spin'); majBadge(); majEtat(); });
  }
  function majEtat() {
    var n = state.rapports.filter(function (r) { return r._etat === 'a_envoyer'; }).length, e = $('#etat'), R = state.reglages;
    var t = !R.url || !R.cle ? 'Connexion à Google Sheets non configurée (menu ⋮)' :
      state.sync ? 'Synchronisation…' :
      (n ? n + ' rapport' + (n > 1 ? 's' : '') + ' en attente d’envoi · ' : '') + (state.derniere ? 'synchronisé à ' + new Date(state.derniere).toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' }) : 'jamais synchronisé');
    if (state.erreur && !state.sync) t = state.erreur + (n ? ' (' + n + ' en attente)' : '');
    e.textContent = t; e.classList.toggle('warn', !!(n || state.erreur || !R.url));
  }
  function majBadge() {
    var n = state.rapports.filter(function (r) { return r._etat === 'a_envoyer' || r._etat === 'brouillon'; }).length, b = $('#dot-attente');
    b.hidden = !n; b.textContent = n;
  }

  /* ---------- Mes rapports ---------- */
  function rendreListe() {
    var fp = $('#l-pilier').value, fz = $('#l-zone').value, qui = $('#l-qui').value;
    var L = state.rapports.filter(function (r) {
      return (!fp || r.pilier === fp) && (!fz || r.zone === fz) && (qui === 'tous' || r._local) && !(r.donnees && r.donnees.supprime);
    }).sort(function (a, b) { return b.date.localeCompare(a.date) || String(b.maj).localeCompare(String(a.maj)); });
    $('#l-resume').textContent = L.length + ' rapport' + (L.length > 1 ? 's' : '') + (L.length > 200 ? ' (200 plus récents affichés)' : '');
    var host = $('#l-items');
    host.replaceChildren.apply(host, L.slice(0, 200).map(function (r) {
      var p = D.pilier(r.pilier) || { nom: r.pilier, couleur: '#888' }, e = (r.donnees && r.donnees.entete) || {};
      var unite = Object.keys(e).map(function (k) { return e[k]; }).filter(Boolean).join(', ');
      var badge = r._etat === 'brouillon' ? el('span', { cls: 'badge mute', text: 'Brouillon' }) : r._etat === 'a_envoyer' ? el('span', { cls: 'badge warn', text: 'À envoyer' }) : el('span', { cls: 'badge ok', text: 'Envoyé' });
      var tag = el('span', { cls: 'pilier-tag' }, [el('i'), document.createTextNode(p.court || p.nom)]); tag.firstChild.style.background = p.couleur;
      return el('div', { cls: 'ritem', onclick: function () { ouvrirRapport(r); } }, [
        el('div', { cls: 't' }, [tag, el('b', { text: A.fr(r.date) + ' · ' + r.zone + (unite ? ' · ' + unite : '') }), el('span', { text: (r.auteur || 'auteur inconnu') })]), badge]);
    }));
  }
  function ouvrirRapport(r) {
    if (state.sale && !confirm('Des valeurs saisies ne sont pas enregistrées. Les abandonner ?')) return;
    $('#f-pilier').value = r.pilier; $('#f-date').value = r.date;
    state._pilierAff = null; state.courant = r; state.sale = false;
    var p = D.pilier(r.pilier); construireEntete(p, r.donnees && r.donnees.entete); state._pilierAff = p.id;
    remplir($('#f-zone'), D.zonesPour(p), r.zone);
    chargerSelection(true); montrer('saisir');
  }

  /* ---------- Tableau de bord ---------- */
  function filtresTableau() {
    var fin = $('#t-date').value || A.jour(), per = $('#t-periode').value;
    var debut = per === 'tout' ? '2000-01-01' : per === '0' ? fin : A.decaler(fin, -(+per) + 1);
    var z = $('#t-zone').value;
    return { debut: debut, fin: fin, zones: z ? [z] : [], zone: z, pilier: $('#t-pilier').value, periode: per };
  }
  function rendreTableau() {
    var f = filtresTableau(), R = state.rapports;
    $('#t-info').textContent = (f.periode === '0' ? 'Rapports du ' + A.fr(f.fin) : f.periode === 'tout' ? 'Cumul jusqu’au ' + A.fr(f.fin) : 'Du ' + A.fr(f.debut) + ' au ' + A.fr(f.fin)) +
      ' · ' + (f.zone || 'toutes les zones') + '. Les valeurs marquées « fin de journée » prennent le dernier jour, les autres s’additionnent.';
    // Complétude
    var C = A.completude(R, f.fin), cols = D.ZONES.concat([D.PROVINCE]);
    var tb = $('#t-comp'); tb.replaceChildren();
    var hd = el('tr', {}, [el('th', { cls: 'l', text: 'Pilier' })]);
    cols.forEach(function (z) { hd.append(el('th', { text: z === D.PROVINCE ? 'Prov.' : z })); }); hd.append(el('th', { text: 'ZS' }));
    tb.append(el('thead', {}, [hd]));
    var body = el('tbody');
    D.PILIERS.forEach(function (p) {
      var c = C[p.id] || {}, n = D.ZONES.filter(function (z) { return c[z]; }).length;
      var tr = el('tr', {}, [el('td', { cls: 'l', text: p.court })]);
      cols.forEach(function (z) { tr.append(el('td', { cls: c[z] ? 'ok' : 'no', text: c[z] ? (c[z] > 1 ? '✓' + c[z] : '✓') : '·' })); });
      tr.append(el('td', { text: n + '/13' })); body.append(tr);
    });
    tb.append(body);
    $('#t-comp-sub').textContent = 'Rapports reçus pour le ' + A.fr(f.fin) + ' (✓2 = deux unités, par exemple deux CTE).';
    // Piliers
    var host = $('#t-piliers'); host.replaceChildren();
    D.PILIERS.filter(function (p) { return !f.pilier || p.id === f.pilier; }).forEach(function (p) {
      var c = A.consolider(R, p, f); if (!c.ordre.length) return;
      host.append(blocPilier(p, c, f));
    });
    if (!host.children.length) host.append(el('p', { cls: 'card', text: 'Aucun rapport pour ces filtres. Synchronisez (bouton ⟳) pour recevoir les rapports des autres zones.' }));
  }
  function blocPilier(p, c, f) {
    var box = el('div', { cls: 'card pbloc' }), ic = el('i'); ic.style.background = p.couleur;
    box.append(el('h2', {}, [ic, document.createTextNode(p.nom)]));
    box.append(el('p', { cls: 'sub', text: c.ordre.length + ' zone' + (c.ordre.length > 1 ? 's' : '') + ' · ' + c.total.rapports.length + ' rapport' + (c.total.rapports.length > 1 ? 's' : '') }));
    var cols = A.colonnesCles(p), Ktot = A.calculer(p, c.total.valeurs);
    var val = function (V, K, col) { return fmt(A.valCol(col, V, K), col.pct); };
    box.append(el('div', { cls: 'kpis' }, cols.slice(0, 8).map(function (col) {
      return el('div', { cls: 'kpi' }, [el('span', { text: col.lib }), el('b', { text: val(c.total.valeurs, Ktot, col) })]);
    })));
    // Tableau par zone
    var t = el('table', { cls: 't' }), hd = el('tr', {}, [el('th', { cls: 'l', text: 'Zone' })]);
    cols.forEach(function (col) { hd.append(el('th', { text: col.lib })); });
    t.append(el('thead', {}, [hd]));
    var bd = el('tbody');
    c.ordre.forEach(function (z) {
      var V = c.zones[z].valeurs, K = A.calculer(p, V), tr = el('tr', {}, [el('td', { cls: 'l', text: z })]);
      cols.forEach(function (col) { tr.append(el('td', { text: val(V, K, col) })); }); bd.append(tr);
    });
    if (c.ordre.length > 1) {
      var tr = el('tr', { cls: 'tot' }, [el('td', { cls: 'l', text: 'Total' })]);
      cols.forEach(function (col) { tr.append(el('td', { text: val(c.total.valeurs, Ktot, col) })); }); bd.append(tr);
    }
    t.append(bd); box.append(el('div', { cls: 'tablewrap' }, [t]));
    // Détail complet (repliable)
    var det = el('details', { cls: 'sec', style: 'margin-top:10px' }), sm = el('summary', { text: 'Détail de tous les indicateurs' });
    det.append(sm);
    det.addEventListener('toggle', function () { if (det.open && det.children.length === 1) det.append(detailPilier(p, c)); });
    box.append(det);
    // Textes
    [['saillants', 'Points saillants'], ['realisees', 'Activités réalisées'], ['planifiees', 'Activités planifiées']].forEach(function (x) {
      var T = A.textes(c.total.rapports, x[0]); if (!T.length) return;
      box.append(el('h2', { text: x[1], style: 'margin-top:12px;font-size:14px' }));
      box.append(el('ul', { cls: 'txts' }, T.map(function (t) { return el('li', {}, [el('b', { text: t.zone + (f.debut !== f.fin ? ' (' + A.fr(t.date) + ')' : '') + ' : ' }), document.createTextNode(t.t)]); })));
    });
    var Df = A.lignesListe(c.total.rapports, 'defis');
    if (Df.length) {
      box.append(el('h2', { text: 'Défis', style: 'margin-top:12px;font-size:14px' }));
      box.append(el('ul', { cls: 'txts' }, Df.map(function (x) {
        return el('li', {}, [el('b', { text: x.zone + ' : ' }), document.createTextNode(x.l.defi || ''), x.l.solution ? document.createTextNode(' → ' + x.l.solution) : null,
          x.l.resp ? document.createTextNode(' (' + x.l.resp + (x.l.delai ? ', ' + x.l.delai : '') + ')') : null]);
      })));
    }
    return box;
  }
  function detailPilier(p, c) {
    var t = el('table', { cls: 't' }), hd = el('tr', {}, [el('th', { cls: 'l', text: 'Indicateur' })]);
    c.ordre.forEach(function (z) { hd.append(el('th', { text: z })); }); if (c.ordre.length > 1) hd.append(el('th', { text: 'Total' }));
    t.append(el('thead', {}, [hd]));
    var bd = el('tbody'), Ks = {};
    c.ordre.forEach(function (z) { Ks[z] = A.calculer(p, c.zones[z].valeurs); }); var Kt = A.calculer(p, c.total.valeurs);
    var sec0 = null;
    D.champs(p).forEach(function (ch) {
      if (ch.sec !== sec0) { sec0 = ch.sec; var trh = el('tr', {}, [el('td', { cls: 'l', text: ch.sec.titre, style: 'font-weight:600;color:var(--muted)' })]); bd.append(trh); }
      var tr = el('tr', {}, [el('td', { cls: 'l', text: ch.lib + (ch.agg === 'stock' ? ' *' : '') })]);
      c.ordre.forEach(function (z) { tr.append(el('td', { text: fmt(c.zones[z].valeurs[ch.id]) })); });
      if (c.ordre.length > 1) tr.append(el('td', { text: fmt(c.total.valeurs[ch.id]) }));
      bd.append(tr);
    });
    if ((p.calculs || []).length) {
      bd.append(el('tr', {}, [el('td', { cls: 'l', text: 'Indicateurs calculés', style: 'font-weight:600;color:var(--muted)' })]));
      p.calculs.forEach(function (k) {
        var tr = el('tr', {}, [el('td', { cls: 'l', text: k.lib })]);
        c.ordre.forEach(function (z) { tr.append(el('td', { text: fmt(Ks[z][k.id], k.pct) })); });
        if (c.ordre.length > 1) tr.append(el('td', { text: fmt(Kt[k.id], k.pct) }));
        bd.append(tr);
      });
    }
    t.append(bd);
    return el('div', { cls: 'tablewrap', style: 'margin:0 12px 12px' }, [t, el('p', { cls: 'note', style: 'padding:0 8px 8px', text: '* Situation en fin de journée : dernier jour de la période.' })]);
  }

  /* ---------- Exports ---------- */
  function exporter(type) {
    var f = filtresTableau();
    if (type === 'xlsx') {
      busy(true, 'Préparation du classeur…');
      setTimeout(function () {
        try { window.EXPORTS.excel(state.rapports, f); toast('Classeur Excel enregistré dans les téléchargements.'); }
        catch (e) { toast('Export impossible : ' + e.message, 6000); }
        busy(false);
      }, 30);
      return;
    }
    busy(true, 'Préparation des diapositives…');
    chargerScript('pptxgen.bundle.js', function () { return window.PptxGenJS; }).then(function () {
      return window.EXPORTS.diapos(state.rapports, f);
    }).then(function () { toast('Présentation enregistrée dans les téléchargements.'); })
      .catch(function (e) { toast('Diapositives impossibles : ' + e.message, 6000); })
      .then(function () { busy(false); });
  }
  function chargerScript(src, pret) {
    if (pret()) return Promise.resolve();
    return new Promise(function (res, rej) {
      var s = document.createElement('script'); s.src = src; s.onload = function () { res(); };
      s.onerror = function () { rej(new Error('fichier ' + src + ' introuvable (ouvrez l’application une fois avec Internet)')); };
      document.head.append(s);
    });
  }

  /* ---------- Vues, réglages ---------- */
  function montrer(v) {
    state.vue = v; lsSet('piliers-vue', v);
    ['saisir', 'liste', 'qualite', 'tableau'].forEach(function (x) { $('#v-' + x).hidden = x !== v; });
    $$('.tabs button').forEach(function (b) { b.setAttribute('aria-selected', b.dataset.v === v ? 'true' : 'false'); });
    window.scrollTo(0, 0); rafraichir();
  }
  function rafraichir() { if (state.vue === 'liste') rendreListe(); if (state.vue === 'tableau') rendreTableau(); if (state.vue === 'qualite') rendreQualite(); majEtat(); majBadgeQualite(); }

  /* ---------- Qualité ---------- */
  function filtresQualite() {
    var fin = $('#q-date').value || A.jour(), per = $('#q-periode').value, z = $('#q-zone').value;
    return { debut: per === 'tout' ? '2000-01-01' : per === '0' ? fin : A.decaler(fin, -(+per) + 1), fin: fin, zones: z ? [z] : [], zone: z, pilier: $('#q-pilier').value, periode: per };
  }
  function majBadgeQualite() {
    var j = A.jour(), R = Q.inspecter(state.rapports, { debut: j, fin: j, zones: [] }), n = R.erreurs + R.ecarts, b = $('#dot-qualite');
    b.hidden = !n; b.textContent = n;
  }
  function nomPilier(id) { return [].concat(id).map(function (x) { var p = D.pilier(x); return p ? p.court : x; }).join(' + '); }
  function rendreQualite() {
    var f = filtresQualite(), R = Q.inspecter(state.rapports, f);
    var k = $('#q-kpis');
    k.replaceChildren.apply(k, [
      ['Rapports contrôlés', R.rapports, ''], ['Rapports sans anomalie', R.rapports ? pf.format(R.sansAnomalie / R.rapports) : '–', R.rapports && R.sansAnomalie === R.rapports ? 'ok' : 'warn'],
      ['Incohérences (à corriger)', R.erreurs, R.erreurs ? 'bad' : 'ok'], ['Valeurs à vérifier', R.alertes, R.alertes ? 'warn' : 'ok'],
      ['Ruptures avec la veille', R.ruptures, R.ruptures ? 'warn' : 'ok'], ['Écarts entre piliers', R.ecarts, R.ecarts ? 'bad' : 'ok'],
      ['Concordances respectées', R.concOk, 'ok'], ['Comparaisons impossibles (rapport manquant)', R.manques, R.manques ? 'warn' : '']
    ].map(function (x) { return el('div', { cls: 'kpi ' + x[2] }, [el('span', { text: x[0] }), el('b', { text: typeof x[1] === 'number' ? nf.format(x[1]) : x[1] })]); }));
    // Concordance
    var tout = $('#q-tout').checked, host = $('#q-conc'); host.replaceChildren();
    $('#q-conc-sub').textContent = 'Pour chaque zone, les chiffres de deux piliers qui doivent concorder (' + (f.periode === '0' ? 'le ' + A.fr(f.fin) : 'du ' + A.fr(f.debut) + ' au ' + A.fr(f.fin)) +
      '). Sur un seul jour, un décalage de saisie peut créer un écart : comparez aussi sur 7 jours.';
    Q.CONCORDANCES.forEach(function (Rg) {
      var L = R.concordances.filter(function (x) { return x.regle === Rg && (tout || x.statut !== 'ok'); });
      if (!L.length) return;
      var t = el('table', { cls: 't conc' });
      t.append(el('thead', {}, [el('tr', {}, [el('th', { cls: 'l', text: 'Zone' }), el('th', { text: '①' }), el('th', { text: '②' }), el('th', { text: 'Statut' })])]));
      t.append(el('tbody', {}, L.map(function (x) {
        var e = x.a != null && x.b != null ? x.a - x.b : null;
        var st = x.statut === 'ok' ? ['ok', 'Concordant'] : x.statut === 'ecart' ? ['ecart', 'Écart ' + (e > 0 ? '+' : '') + Q.nb(e)] : ['manque', 'Sans rapport : ' + nomPilier(x.manque)];
        return el('tr', { cls: x.zone === 'Total' ? 'tot' : null }, [el('td', { cls: 'l', text: x.zone === 'Total' ? 'Total' + (x.partiel ? ' (zones comparables)' : '') : x.zone }),
          el('td', { text: Q.nb(x.a) }), el('td', { text: Q.nb(x.b) }), el('td', {}, [el('span', { cls: 'st ' + st[0], text: st[1] })])]);
      })));
      host.append(el('div', { cls: 'regle' }, [el('h3', { text: Rg.lib }),
        el('p', { cls: 'aide' }, [el('b', { text: '① ' }), document.createTextNode(Rg.a.lib + '  '), el('b', { text: (Rg.type === 'eg' ? '=' : '≤') + ' ② ' }), document.createTextNode(Rg.b.lib)]),
        Rg.aide ? el('p', { cls: 'aide', text: Rg.aide }) : null, el('div', { cls: 'tablewrap' }, [t])]));
    });
    if (!host.children.length) host.append(el('p', { cls: 'note', text: R.concordances.length ? '✓ Toutes les concordances vérifiables sont respectées.' : 'Aucune comparaison possible : il faut les rapports de deux piliers pour une même zone.' }));
    // Rapports à corriger
    var lr = $('#q-rap');
    $('#q-rap-sub').textContent = R.aCorriger.length ? R.aCorriger.length + ' rapport' + (R.aCorriger.length > 1 ? 's' : '') + ' avec au moins une anomalie. Touchez un rapport pour l’ouvrir et le corriger.' : '';
    lr.replaceChildren.apply(lr, R.aCorriger.length ? R.aCorriger.map(function (x) {
      var p = D.pilier(x.r.pilier) || { court: x.r.pilier, couleur: '#888' }, tag = el('span', { cls: 'pilier-tag' }, [el('i'), document.createTextNode(p.court)]); tag.firstChild.style.background = p.couleur;
      var msgs = el('ul', { cls: 'ctrl', style: 'margin-top:6px' }, x.internes.map(function (i) {
        return el('li', { cls: i.regle.grav }, [el('b', { text: i.regle.lib }), i.a != null || i.b != null ? el('small', { text: 'Valeurs : ' + Q.nb(i.a) + ' / ' + Q.nb(i.b) }) : null]);
      }).concat(x.continuite.map(function (c) { return el('li', { cls: 'alerte' }, [el('b', { text: c.regle.lib }), el('small', { text: 'Saisi : ' + Q.nb(c.a) + ' · veille : ' + Q.nb(c.b) })]); })));
      var e = (x.r.donnees && x.r.donnees.entete) || {}, u = Object.keys(e).map(function (kk) { return e[kk]; }).filter(Boolean).join(', ');
      return el('div', { cls: 'card', style: 'cursor:pointer;padding:10px 12px', onclick: function () { ouvrirRapport(x.r); } },
        [tag, el('b', { text: ' ' + A.fr(x.r.date) + ' · ' + x.r.zone + (u ? ' · ' + u : '') }), el('div', { cls: 'note', style: 'margin:2px 0 0', text: x.r.auteur || '' }), msgs]);
    }) : [el('p', { cls: 'note', text: R.rapports ? '✓ Aucun rapport incohérent sur cette période.' : 'Aucun rapport sur cette période.' })]);
  }
  function rendreRegles() {
    var h = $('#q-regles'); h.replaceChildren();
    h.append(el('h3', { text: 'Cohérence de chaque rapport', style: 'font-size:14px;margin:4px 0' }));
    D.PILIERS.forEach(function (p) { (Q.INTERNES[p.id] || []).forEach(function (r) { h.append(el('p', { cls: 'note', text: p.court + ' · ' + (r.grav === 'erreur' ? 'incohérence' : 'à vérifier') + ' : ' + r.lib })); }); });
    h.append(el('h3', { text: 'Continuité avec la veille', style: 'font-size:14px;margin:10px 0 4px' }));
    Object.keys(Q.CONTINUITE).forEach(function (k) { Q.CONTINUITE[k].forEach(function (r) { h.append(el('p', { cls: 'note', text: D.pilier(k).court + ' : ' + r.lib })); }); });
    h.append(el('h3', { text: 'Concordance entre piliers', style: 'font-size:14px;margin:10px 0 4px' }));
    Q.CONCORDANCES.forEach(function (r) { h.append(el('p', { cls: 'note', text: nomPilier(r.a.pilier) + ' ↔ ' + nomPilier(r.b.pilier) + ' : ' + r.lib })); });
  }
  function ouvrirReglages() {
    var R = state.reglages;
    $('#r-auteur').value = R.auteur || ''; $('#r-url').value = R.url || ''; $('#r-cle').value = R.cle || '';
    remplir($('#r-pilier'), listePiliers, R.pilier || '', '—'); remplir($('#r-zone'), toutesZones, R.zone || '', '—');
    $('#r-msg').textContent = '';
    $('#sheet-bg').hidden = false; $('#sheet-reglages').hidden = false;
  }
  function fermer() { $('#sheet-bg').hidden = true; $('#sheet-reglages').hidden = true; }
  function sauverReglages() {
    var R = { auteur: $('#r-auteur').value.trim(), url: $('#r-url').value.trim(), cle: $('#r-cle').value, pilier: $('#r-pilier').value, zone: $('#r-zone').value };
    if (R.url && !/^https:\/\/script\.google(usercontent)?\.com\//.test(R.url)) { $('#r-msg').textContent = 'L’adresse doit commencer par https://script.google.com/'; return; }
    state.reglages = R; S.set('reglages', R);
    if (!R.url || !R.cle) { $('#r-msg').textContent = 'Réglages enregistrés. Sans adresse ni clé, les rapports restent sur ce téléphone.'; return; }
    $('#r-msg').textContent = 'Test de la connexion…';
    appel({ action: 'lire', depuis: '9999' }).then(function () {
      $('#r-msg').textContent = 'Connexion réussie.'; state.erreur = null; synchroniser(false);
    }).catch(function (e) { $('#r-msg').textContent = e.message; });
  }

  /* ---------- Démarrage ---------- */
  function init() {
    remplir($('#f-pilier'), listePiliers); remplir($('#l-pilier'), listePiliers, '', 'Tous les piliers');
    remplir($('#l-zone'), toutesZones, '', 'Toutes les zones'); remplir($('#t-zone'), toutesZones, '', 'Toutes les zones');
    remplir($('#t-pilier'), listePiliers, '', 'Tous les piliers');
    $('#f-date').value = A.jour(); $('#t-date').value = A.jour();
    ['#f-pilier', '#f-date', '#f-zone'].forEach(function (s) { $(s).addEventListener('change', function () { chargerSelection(); }); });
    $('#btn-brouillon').addEventListener('click', function () { enregistrer(false); });
    $('#btn-envoyer').addEventListener('click', function () { enregistrer(true); });
    $('#btn-supprimer').addEventListener('click', supprimer);
    $('#btn-nouveau').addEventListener('click', function () {
      if (state.sale && !confirm('Abandonner les valeurs non enregistrées ?')) return;
      $('#f-date').value = A.jour(); state.sale = false; chargerSelection(true); window.scrollTo(0, 0);
    });
    $$('.tabs button').forEach(function (b) { b.addEventListener('click', function () { montrer(b.dataset.v); }); });
    ['#l-pilier', '#l-zone', '#l-qui'].forEach(function (s) { $(s).addEventListener('change', rendreListe); });
    ['#t-date', '#t-periode', '#t-zone', '#t-pilier'].forEach(function (s) { $(s).addEventListener('change', rendreTableau); });
    remplir($('#q-zone'), D.ZONES, '', 'Toutes les zones'); remplir($('#q-pilier'), listePiliers, '', 'Tous les piliers'); $('#q-date').value = A.jour();
    ['#q-date', '#q-periode', '#q-zone', '#q-pilier', '#q-tout'].forEach(function (s) { $(s).addEventListener('change', rendreQualite); });
    rendreRegles();
    $('#btn-xlsx').addEventListener('click', function () { exporter('xlsx'); });
    $('#btn-pptx').addEventListener('click', function () { exporter('pptx'); });
    $('#btn-sync').addEventListener('click', function () { synchroniser(false); });
    $('#btn-reglages').addEventListener('click', ouvrirReglages);
    $('#sheet-bg').addEventListener('click', fermer); $$('[data-close]').forEach(function (b) { b.addEventListener('click', fermer); });
    $('#btn-r-ok').addEventListener('click', sauverReglages);
    $('#btn-r-tout').addEventListener('click', function () { fermer(); synchroniser(false, true); });
    $('#r-voir').addEventListener('change', function () { $('#r-cle').type = $('#r-voir').checked ? 'text' : 'password'; });
    window.addEventListener('online', function () { state.erreur = null; synchroniser(true); });
    window.addEventListener('beforeunload', function (e) { if (state.sale) { e.preventDefault(); e.returnValue = ''; } });
    document.addEventListener('visibilitychange', function () {
      if (!document.hidden && (!state.derniere || Date.now() - state.derniere > 120000)) synchroniser(true);
    });
    setInterval(function () { if (!document.hidden) synchroniser(true); }, 5 * 60000);

    Promise.all([S.get('reglages'), S.tous(), S.get('derniereSync')]).then(function (res) {
      state.reglages = res[0] || {}; state.rapports = res[1] || []; state.derniere = res[2] || null;
      if (state.reglages.pilier) $('#f-pilier').value = state.reglages.pilier;
      chargerSelection(true); majBadge();
      montrer(lsGet('piliers-vue') || 'saisir');
      if (!state.reglages.auteur) ouvrirReglages(); else synchroniser(true);
    }).catch(function () { toast('Stockage du téléphone indisponible : utilisez Chrome hors navigation privée.', 8000); });
    if ('serviceWorker' in navigator && location.protocol !== 'file:') navigator.serviceWorker.register('sw.js').catch(function () { /* hors ligne indisponible */ });
  }
  init();
})();
