/* ============================================================================
   SNAF Cotes — COUCHE ONEDRIVE (Microsoft Graph) — v2
   ----------------------------------------------------------------------------
   Remplace la couche serveur (server.py) de la PWA par un stockage direct
   dans le OneDrive partagé « CHANTIERS EN COURS » (Kanban par statut).

   Connexion : chaque utilisateur avec SON compte Microsoft (OAuth MSAL).
   Rôles     : mappés par e-mail -> BB / NR / MP (règle Bastien 20/08/2026).
   Données   : « CHANTIERS EN COURS/<Statut>/<Client - Lieu>/releve.json »
               + photos/PDF dans le même dossier.
   Conflits  : champ _rev dans releve.json ; PUT périmé -> erreur 409 + modale.

   Chargé APRÈS le script principal ; surcharge les fonctions biblio*.
   La production (snaf-cotes-server) n'est PAS concernée.
   ============================================================================ */
(function () {
  'use strict';

  var CONFIG = {
    clientId: 'dfd1ff8e-0f80-4cee-b2d6-a3b696008b31',
    tenant: '3d854a15-694a-4818-b60b-a134457b63d4',
    driveId: 'b!TPSmx6V9fUG5DiB27ODQajb0UpCc3LxKuyGg1gvo1QTzF6UPkQ-VR7dO0CV-PPKY',
    racineId: '015BFCG4PCD6EQ5R2ZN5H2GDYQXTBKVIFS',
    scopes: ['Files.ReadWrite', 'User.Read'],
    users: [
      { emails: ['bastien.blanc@mecaservice.fr', 'bastien.blanc@snaf83.fr', 'bastien.blanc13@gmail.com'], initials: 'BB', role: 'admin', canDelete: true },
      { emails: ['nicolas.racenet@snaf83.fr'], initials: 'NR', role: 'NR', canDelete: false },
      { emails: ['accueil@snaf83.fr'], initials: 'MP', role: 'MP', canDelete: false }
    ]
  };

  var MAP_STATUT = { brouillon: 'Brouillon', devis: 'Devis', commande: 'Commande', fabrication: 'Fabrication', pose: 'Installé' };
  var STATUT_CLE = { Brouillon: 'brouillon', Devis: 'devis', Commande: 'commande', Fabrication: 'fabrication', 'Installé': 'pose' };
  var NR_STATUTS_OK = { commande: 1, fabrication: 1, pose: 1 };
  var GRAPH = 'https://graph.microsoft.com/v1.0';

  var _msal = null, _account = null, _token = null, _tokenExp = 0;
  var _redirectHandled = false;
  // Cache d'index : byId[rid] = {id, folderId, name, statutDir, statut, data}
  var _cache = { byId: {}, statutDirs: {}, ts: 0 };
  var CACHE_TTL = 60000; // 60 s

  /* ------------------------------ OUTILS ----------------------------------- */

  function _cleanName(s) {
    s = String(s || '').replace(/[\\/:*?"<>|]/g, '-').replace(/\s+/g, ' ').trim();
    return s || null;
  }
  function _nomDossier(client, lieu) {
    client = String(client || '').trim();
    lieu = String(lieu || '').trim();
    if (!client) return null;
    var bas = (lieu && lieu.toLowerCase() !== 'a definir' && lieu.toLowerCase() !== 'à définir')
      ? client + ' - ' + lieu : client;
    return _cleanName(bas);
  }
  function _lower(s) { return String(s || '').toLowerCase().trim(); }
  function _err(status, msg, data) {
    var e = new Error(msg || ('Erreur ' + status));
    e.status = status;
    if (data !== undefined) e.data = data;
    return e;
  }
  function _b64ToU8(b64) {
    var bin = atob(b64);
    var u8 = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
    return u8;
  }
  function _textToU8(txt) {
    if (typeof TextEncoder !== 'undefined') return new TextEncoder().encode(txt);
    var b64 = btoa(unescape(encodeURIComponent(txt)));
    return _b64ToU8(b64);
  }
  function _parseJsonDoc(txt) {
    if (!txt) return null;
    if (typeof txt === 'object') return txt;
    try { return JSON.parse(txt); } catch (e) { return null; }
  }
  function _clone(o) { return JSON.parse(JSON.stringify(o)); }

  // Requête Graph authentifiée. binaire=true -> body envoyé tel quel, réponse en texte.
  function _graph(method, path, body, isBinary) {
    if (!_token) return Promise.reject(_err(401, 'Non connecté'));
    var headers = { Authorization: 'Bearer ' + _token };
    var opts = { method: method, headers: headers };
    if (body !== undefined && body !== null) {
      if (isBinary) {
        opts.body = (typeof body === 'string') ? body : body;
        headers['Content-Type'] = 'application/octet-stream';
      } else {
        opts.body = typeof body === 'string' ? body : JSON.stringify(body);
        headers['Content-Type'] = 'application/json';
      }
    }
    return fetch(GRAPH + path, opts).then(function (r) {
      return r.text().then(function (txt) {
        var data = null;
        var ct = (r.headers.get('content-type') || '').toLowerCase();
        if (ct.indexOf('json') >= 0) { try { data = JSON.parse(txt); } catch (e) { data = txt; } }
        else if (r.status === 204) data = {};
        else data = txt;
        if (!r.ok) {
          var msg = (data && data.error && (data.error.message || data.error.code)) || ('Erreur Graph ' + r.status);
          var e = _err(r.status === 404 ? 404 : r.status, String(msg).slice(0, 200));
          e.raw = data;
          throw e;
        }
        return data;
      });
    }).catch(function (e) {
      if (e && e.status) throw e;
      throw _err(0, 'Réseau injoignable');
    });
  }
  function _item(idOrPath) { return '/drives/' + CONFIG.driveId + '/items/' + idOrPath; }

  /* --------------------------- AUTH MSAL ----------------------------------- */

  function _redirectUri() { return window.location.origin + window.location.pathname; }
  function _msalInit() {
    if (_msal) return _msal;
    if (typeof msal === 'undefined' || !msal.PublicClientApplication) return null;
    _msal = new msal.PublicClientApplication({
      auth: {
        clientId: CONFIG.clientId,
        authority: 'https://login.microsoftonline.com/' + CONFIG.tenant,
        redirectUri: _redirectUri()
      },
      cache: { cacheLocation: 'localStorage', storeAuthStateInCookie: true }
    });
    return _msal;
  }

  function _roleFor(email) {
    email = _lower(email);
    for (var i = 0; i < CONFIG.users.length; i++) {
      var u = CONFIG.users[i];
      for (var j = 0; j < u.emails.length; j++) {
        if (_lower(u.emails[j]) === email) return u;
      }
    }
    return null;
  }

  function _appliquerCompte(account, token, expiresIn) {
    _account = account;
    _token = token;
    _tokenExp = Date.now() + (expiresIn || 3600) * 1000;
    biblio.code = null;
    try { localStorage.removeItem('snaf_biblio_code'); } catch (e) {}
    var role = _roleFor(account.username);
    if (!role) role = { initials: '?', role: 'unknown', canDelete: false };
    biblio.user = {
      initials: role.initials,
      role: role.role,
      canDelete: role.canDelete,
      email: account.username,
      displayName: account.name || account.username
    };
    return role;
  }

  function _ensureToken() {
    if (!_msalInit()) return Promise.reject(_err(0, 'MSAL non chargé'));
    if (_token && Date.now() < _tokenExp - 60000) return Promise.resolve(_token);
    var acc = _account || (_msal.getAllAccounts()[0] || null);
    if (!acc) return Promise.reject(_err(401, 'Aucun compte connecté'));
    _account = acc;
    return _msal.acquireTokenSilent({ scopes: CONFIG.scopes, account: acc })
      .then(function (res) {
        _token = res.accessToken; _tokenExp = Date.now() + (res.expiresIn || 3600) * 1000;
        return _token;
      })
      .catch(function () {
        return _msal.acquireTokenPopup({ scopes: CONFIG.scopes, account: acc }).then(function (res) {
          _token = res.accessToken; _tokenExp = Date.now() + (res.expiresIn || 3600) * 1000;
          return _token;
        });
      });
  }

  function msalLogin() {
    var inst = _msalInit();
    if (!inst) { if (typeof biblioToast === 'function') biblioToast('Bibliothèque de connexion absente'); return; }
    var acc = inst.getAllAccounts()[0];
    if (acc) {
      _account = acc;
      _ensureToken().then(function () { return _connecter(acc); })
        .catch(function () { deconnecter(); });
      return;
    }
    inst.loginRedirect({ scopes: CONFIG.scopes });
  }

  // Traite le retour de redirection OAuth (une seule fois)
  // Restaure d'abord une session existante (cache localStorage MSAL) : quand on
  // rouvre l'appli normalement, on reconnecte immédiatement SANS popup ni écran
  // de login. handleRedirectPromise ne sert que pour le vrai retour OAuth.
  function _handleRedirect() {
    if (_redirectHandled) return;
    var inst = _msalInit();
    if (!inst) return;
    _redirectHandled = true;
    var acc = inst.getAllAccounts()[0];
    if (acc && !window.location.hash) {
      // Session en cache : restauration silencieuse (pas de popup bloquante)
      _account = acc;
      _ensureToken().then(function () {
        return _connecter(acc);
      }).catch(function () {
        // Token silencieux impossible -> on laisse handleRedirectPromise tenter,
        // sinon écran de connexion.
        inst.handleRedirectPromise().then(function (res) {
          if (res && res.account) {
            _appliquerCompte(res.account, res.accessToken, res.expiresIn);
            return _connecter(res.account);
          }
          afficherLogin();
          return null;
        }).catch(function () {
          afficherLogin();
        });
      });
      return;
    }
    inst.handleRedirectPromise().then(function (res) {
      if (res && res.account) {
        _appliquerCompte(res.account, res.accessToken, res.expiresIn);
        return _connecter(res.account);
      }
      var acc2 = inst.getAllAccounts()[0];
      if (acc2) {
        _account = acc2;
        return _ensureToken().then(function () { return _connecter(acc2); });
      }
      afficherLogin();
      return null;
    }).catch(function () {
      afficherLogin();
    });
  }

  function _connecter(account) {
    var role = _appliquerCompte(account, _token, 3600);
    if (role.role === 'unknown') {
      if (typeof biblioToast === 'function') biblioToast('Ce compte n\'est pas autorisé pour SNAF Cotes');
      deconnecter();
      return Promise.resolve();
    }
    return biblioCharger();
  }

  function deconnecter() {
    _account = null; _token = null; _tokenExp = 0;
    _cache = { byId: {}, statutDirs: {}, ts: 0 };
    biblio.code = null; biblio.on = false; biblio._rev = null; biblio.user = null;
    try { localStorage.removeItem('snaf_biblio_code'); } catch (e) {}
    var inst = _msalInit();
    if (inst) {
      var accs = inst.getAllAccounts();
      if (accs.length) {
        try { inst.logoutRedirect({ account: accs[0] }); return; } catch (e) {
          try { inst.logoutPopup({ account: accs[0] }); } catch (e2) {}
        }
      }
    }
    afficherLogin();
  }

  function afficherLogin() {
    if (typeof biblioAfficher === 'function') biblioAfficher('login');
    var btn = document.getElementById('msLoginBtn');
    if (btn) { btn.style.display = ''; btn.disabled = false; }
  }

  /* --------------------------- INDEX (scan OneDrive) ----------------------- */

  function _statutDirs(force) {
    if (!force && _cache.statutDirs.Brouillon && Date.now() - _cache.ts < CACHE_TTL) {
      return Promise.resolve(_cache.statutDirs);
    }
    return _graph('GET', _item(CONFIG.racineId) + '/children?$select=id,name,folder').then(function (d) {
      var dirs = {};
      ((d && d.value) || []).forEach(function (c) {
        if (c.folder && STATUT_CLE[c.name]) dirs[c.name] = c.id;
      });
      var manquants = Object.keys(MAP_STATUT).map(function (k) { return MAP_STATUT[k]; })
        .filter(function (n) { return !dirs[n]; });
      var p = Promise.resolve();
      manquants.forEach(function (n) {
        p = p.then(function () {
          return _graph('POST', _item(CONFIG.racineId) + '/children', { name: n, folder: {} })
            .then(function (c) { dirs[n] = c.id; })
            .catch(function () {});
        });
      });
      return p.then(function () {
        _cache.statutDirs = dirs;
        return dirs;
      });
    });
  }

  function _lireDossierInfo(folderId, dossierName, statutDir, statutCle) {
    return _graph('GET', _item(folderId) + ':/releve.json:/content').then(function (txt) {
      var data = _parseJsonDoc(txt);
      if (!data || !data.id) return null; // dossier non-SNAF (ex: collègues)
      var id = data.id;
      var statut = data.statut || statutCle || 'brouillon';
      return { id: id, folderId: folderId, name: dossierName, statutDir: statutDir, statut: statut, data: data };
    }).catch(function () { return null; });
  }

  // Scan complet : parcourt les 5 sous-dossiers, lit chaque releve.json.
  function _scan(force) {
    if (!force && Object.keys(_cache.byId).length && Date.now() - _cache.ts < CACHE_TTL) {
      return Promise.resolve(_cache.byId);
    }
    return _statutDirs(force).then(function (dirs) {
      var byId = {};
      var p = Promise.resolve();
      Object.keys(dirs).forEach(function (statutNom) {
        var dirId = dirs[statutNom];
        var statutCle = STATUT_CLE[statutNom];
        p = p.then(function () {
          return _graph('GET', _item(dirId) + '/children?$select=id,name,folder').then(function (d) {
            var folders = ((d && d.value) || []).filter(function (c) { return c.folder; });
            var reads = folders.map(function (c) {
              return _lireDossierInfo(c.id, c.name, statutNom, statutCle).then(function (info) {
                if (info) byId[info.id] = info;
              });
            });
            return Promise.all(reads);
          }).catch(function () { return null; });
        });
      });
      return p.then(function () {
        _cache.byId = byId;
        _cache.ts = Date.now();
        return byId;
      });
    });
  }

  function _findDossier(rid) {
    if (_cache.byId[rid]) return Promise.resolve(_cache.byId[rid]);
    return _scan(false).then(function (byId) { return byId[rid] || null; });
  }

  /* ----------------------------- ROUTES DATA ------------------------------- */

  function _me() {
    if (!biblio.user) return Promise.reject(_err(401, 'Non connecté'));
    return Promise.resolve({
      initials: biblio.user.initials,
      role: biblio.user.role,
      canDelete: !!biblio.user.canDelete,
      email: biblio.user.email
    });
  }

  function _lister() {
    return _scan(false).then(function (byId) {
      var role = (biblio.user && biblio.user.role) || 'BB';
      var out = [];
      Object.keys(byId).forEach(function (rid) {
        var info = byId[rid];
        if (role === 'NR' && !NR_STATUTS_OK[info.statut]) return;
        var d = info.data || {};
        var ch = d.chantier || {};
        var nbPortes = (d.portes || []).length +
          ((d.elements || []).filter(function (e) { return e && e.type && e.type !== 'porte' && e.type !== 'chassis'; }).length);
        out.push({
          id: rid,
          nom: ch.client ? (ch.client + (ch.lieu ? ' - ' + ch.lieu : '')) : info.name,
          client: ch.client || '',
          lieu: ch.lieu || '',
          ref: d.ref || '',
          statut: info.statut || 'brouillon',
          nbPortes: nbPortes,
          modifie: d.dernierDate || 0,
          _folderId: info.folderId
        });
      });
      out.sort(function (a, b) { return String(a.nom).localeCompare(String(b.nom), 'fr'); });
      return { releves: out };
    });
  }

  function _ouvrir(rid) {
    return _findDossier(rid).then(function (info) {
      if (!info) throw _err(404, 'Dossier introuvable');
      var role = (biblio.user && biblio.user.role) || 'BB';
      if (role === 'NR' && !NR_STATUTS_OK[info.statut]) throw _err(403, 'Accès refusé');
      // Relire le JSON à jour (toujours frais au moment de l'ouverture)
      return _graph('GET', _item(info.folderId) + ':/releve.json:/content').then(function (txt) {
        var data = _parseJsonDoc(txt);
        if (!data) throw _err(404, 'releve.json illisible');
        if (!data.id) data.id = rid;
        info.data = data;
        info.statut = data.statut || info.statut;
        return data;
      });
    });
  }

  var _LIBELLES = {
    client: 'client', lieu: 'lieu', adressePose: 'adresse de pose', mailContact: 'email de contact',
    telContact: 'téléphone', notes: 'notes', date: 'date', ref: 'référence', statut: 'statut',
    passL: 'largeur de passage', passH: 'hauteur de passage', baieL: 'largeur de baie', baieH: 'hauteur de baie',
    imposte: 'imposte', fixes: 'fixes', type: 'type', serie: 'série', teinte: 'teinte', vitrage: 'vitrage',
    nbVantaux: 'nombre de vantaux', ouverture: 'sens d\'ouverture'
  };
  function _lib(k) {
    if (_LIBELLES[k]) return _LIBELLES[k];
    return String(k || '').replace(/([A-Z])/g, ' $1').replace(/_/g, ' ').toLowerCase().trim();
  }
  function _diff(local, serveur) {
    var out = [];
    var par = (serveur && serveur.dernierPar) || '?';
    var add = function (ref, champ, l, s) {
      if (out.length >= 10) return;
      var lv = l === undefined ? undefined : (typeof l === 'object' ? JSON.stringify(l) : l);
      var sv = s === undefined ? undefined : (typeof s === 'object' ? JSON.stringify(s) : s);
      if (String(lv) === String(sv)) return;
      out.push({ par: par, ref: ref || '', champ: champ, local: lv === undefined ? null : lv, serveur: sv === undefined ? null : sv });
    };
    var cl = (local && local.chantier) || {}, cs = (serveur && serveur.chantier) || {};
    Object.keys(cl).forEach(function (k) { if (k !== 'frAdresse') add('', _lib(k), cl[k], cs[k]); });
    ['ref', 'statut'].forEach(function (k) { add('', _lib(k), local[k], serveur[k]); });
    var pl = (local && local.portes) || [], ps = (serveur && serveur.portes) || [];
    for (var i = 0; i < Math.max(pl.length, ps.length) && out.length < 10; i++) {
      var a = pl[i] || {}, b = ps[i] || {};
      var ref = a.ref || b.ref || ('Porte ' + (i + 1));
      var cles = {};
      Object.keys(a).forEach(function (k) { cles[k] = 1; });
      Object.keys(b).forEach(function (k) { cles[k] = 1; });
      Object.keys(cles).forEach(function (k) {
        if (k.charAt(0) === '_') return;
        add(ref, _lib(k), a[k], b[k]);
      });
    }
    var el = (local && local.elements) || [], es = (serveur && serveur.elements) || [];
    for (var j = 0; j < Math.max(el.length, es.length) && out.length < 10; j++) {
      var x = el[j] || {}, y = es[j] || {};
      var eref = x.ref || y.ref || ('Équipement ' + (j + 1));
      var xc = x.champs || {}, yc = y.champs || {};
      var keys = {};
      Object.keys(xc).forEach(function (k) { keys[k] = 1; });
      Object.keys(yc).forEach(function (k) { keys[k] = 1; });
      Object.keys(keys).forEach(function (k) { add(eref, _lib(k), xc[k], yc[k]); });
    }
    if (!out.length) out.push({ par: par, ref: '', champ: 'contenu', local: JSON.stringify(local), serveur: JSON.stringify(serveur) });
    return out;
  }

  // Écrit releve.json dans un dossier (retourne la version écrite)
  function _ecrireReleve(folderId, data) {
    var content = JSON.stringify(data, null, 1);
    return _graph('PUT', _item(folderId) + ':/releve.json:/content', _textToU8(content), true)
      .then(function () { return data; });
  }

  function _sauver(rid, data) {
    var role = (biblio.user && biblio.user.role) || 'BB';
    var initials = (biblio.user && biblio.user.initials) || '?';
    var force = !!data._force;
    var statut = data.statut || 'brouillon';
    var statutDir = MAP_STATUT[statut] || 'Devis';

    // Recherche du dossier (cache, sinon scan)
    return _findDossier(rid).then(function (info) {
      var faireSauver = function (folderId, statutDirActuel, estNouveau) {
        // Lire la version distante pour calculer la révision
        var lire = estNouveau ? Promise.resolve(null)
          : _graph('GET', _item(folderId) + ':/releve.json:/content').then(function (txt) {
            var d = _parseJsonDoc(txt);
            if (d && d.id && d.id !== rid) return null; // ne jamais écraser un autre dossier
            return d;
          }).catch(function () { return null; });

        return lire.then(function (distant) {
          // CONFLIT : le client envoie une révision périmée
          if (!force && !estNouveau && data._rev && distant && distant._rev) {
            if (data._rev < distant._rev) {
              throw _err(409, 'conflit', { details: _diff(data, distant) });
            }
          }
          var d2 = _clone(data);
          delete d2._force;
          delete d2._folderId;
          delete d2._statutDir;
          d2.id = rid;
          d2.statut = statut;
          d2._rev = (distant && distant._rev ? distant._rev : 0) + 1;
          d2.dernierPar = initials;
          d2.dernierDate = new Date().toISOString();
          return _ecrireReleve(folderId, d2).then(function () {
            return { d2: d2, statutDirActuel: statutDirActuel };
          });
        });
      };

      if (info) {
        // Existant : possible déplacement si le statut change de sous-dossier
        return faireSauver(info.folderId, info.statutDir, false).then(function (r) {
          var dirCible = MAP_STATUT[r.d2.statut || 'brouillon'] || 'Devis';
          info.data = r.d2;
          info.statut = r.d2.statut;
          if (info.statutDir !== dirCible) {
            return _statutDirs(true).then(function (dirs) {
              var dest = dirs[dirCible] || dirs.Devis;
              return _graph('PATCH', _item(info.folderId), { parentReference: { id: dest } })
                .then(function () { info.statutDir = dirCible; return r.d2; })
                .catch(function () { return r.d2; }); // le JSON est déjà à jour
            });
          }
          return r.d2;
        }).then(function (d2) {
          _cache.ts = 0; // invalider le cache (statut/nom peuvent changer)
          return { _rev: d2._rev, id: rid };
        });
      }

      // NOUVEAU dossier : créer dans le sous-dossier du statut
      return _statutDirs(true).then(function (dirs) {
        var dirId = dirs[statutDir] || dirs.Devis;
        var client = (data.chantier && data.chantier.client) || rid;
        var lieu = (data.chantier && data.chantier.lieu) || '';
        var nom = _nomDossier(client, lieu) || rid;
        // Éviter les doublons : si un dossier du même nom existe déjà dans ce statut
        return _graph('GET', _item(dirId) + '/children?$select=id,name,folder').then(function (lst) {
          var exist = null;
          ((lst && lst.value) || []).forEach(function (c) {
            if (c.folder && c.name === nom) exist = c;
          });
          if (exist) {
            return faireSauver(exist.id, statutDir, true).then(function (r) {
              _cache.ts = 0;
              return { _rev: r.d2._rev, id: rid };
            });
          }
          return _graph('POST', _item(dirId) + '/children', { name: nom, folder: {} }).then(function (c) {
            return faireSauver(c.id, statutDir, true).then(function (r) {
              _cache.ts = 0;
              return { _rev: r.d2._rev, id: rid };
            });
          });
        });
      });
    });
  }

  function _fichiersLister(rid) {
    return _findDossier(rid).then(function (info) {
      if (!info) throw _err(404, 'Dossier introuvable');
      return _graph('GET', _item(info.folderId) + '/children?$select=id,name,size,lastModifiedDateTime').then(function (d) {
        var liste = ((d && d.value) || []).filter(function (c) {
          return !c.folder && c.name !== 'releve.json';
        }).map(function (c) {
          return { nom: c.name, taille: c.size || 0, modifie: c.lastModifiedDateTime || '', id: c.id };
        });
        return { fichiers: liste };
      });
    });
  }

  function _fichierSauver(rid, nom, body) {
    return _findDossier(rid).then(function (info) {
      if (!info) throw _err(404, 'Dossier introuvable');
      var b64 = (body && body.dataBase64) || '';
      if (!b64) throw _err(400, 'Pas de contenu');
      var u8 = _b64ToU8(b64);
      return _graph('PUT', _item(info.folderId) + ':/' + encodeURIComponent(nom) + ':/content', u8, true)
        .then(function () { return { ok: true, nom: nom }; });
    });
  }

  function _fichierBlob(rid, nom) {
    return _findDossier(rid).then(function (info) {
      if (!info) throw _err(404, 'Dossier introuvable');
      return _ensureToken().then(function () {
        var url = GRAPH + _item(info.folderId) + ':/' + encodeURIComponent(nom) + ':/content';
        return fetch(url, { headers: { Authorization: 'Bearer ' + _token } }).then(function (r) {
          if (!r.ok) throw _err(r.status, 'Erreur lecture fichier');
          return r.blob();
        });
      });
    });
  }

  function _supprimer(rid) {
    if (!biblio.user || !biblio.user.canDelete) return Promise.reject(_err(403, 'Accès refusé'));
    return _findDossier(rid).then(function (info) {
      if (!info) throw _err(404, 'Dossier introuvable');
      return _graph('DELETE', _item(info.folderId)).then(function () {
        delete _cache.byId[rid];
        _cache.ts = 0;
        return { ok: true };
      });
    });
  }

  function _changerStatut(rid, statut) {
    return _findDossier(rid).then(function (info) {
      if (!info) throw _err(404, 'Dossier introuvable');
      return _graph('GET', _item(info.folderId) + ':/releve.json:/content').then(function (txt) {
        var data = _parseJsonDoc(txt);
        if (!data) throw _err(404, 'releve.json illisible');
        data.statut = statut;
        data._rev = (data._rev || 0) + 1;
        data.dernierPar = (biblio.user && biblio.user.initials) || '?';
        data.dernierDate = new Date().toISOString();
        return _ecrireReleve(info.folderId, data).then(function () {
          var dirCible = MAP_STATUT[statut] || 'Devis';
          info.data = data;
          info.statut = statut;
          if (info.statutDir !== dirCible) {
            return _statutDirs(true).then(function (dirs) {
              var dest = dirs[dirCible];
              if (!dest) return data;
              return _graph('PATCH', _item(info.folderId), { parentReference: { id: dest } })
                .then(function () { info.statutDir = dirCible; _cache.ts = 0; return data; })
                .catch(function () { _cache.ts = 0; return data; });
            });
          }
          _cache.ts = 0;
          return data;
        });
      });
    });
  }

  /* ------------------------------- ROUTEUR -------------------------------- */

  function biblioApi(url, opts) {
    opts = opts || {};
    var method = (opts.method || 'GET').toUpperCase();
    var m;
    if (url === '/api/me') return _ensureToken().then(_me);
    if (url === '/api/releves' && method === 'GET') return _ensureToken().then(function () { return _lister(); });
    m = url.match(/^\/api\/releves\/([^/]+)$/);
    if (m) {
      var rid = decodeURIComponent(m[1]);
      if (method === 'DELETE') return _ensureToken().then(function () { return _supprimer(rid); });
      if (method === 'PUT') {
        var body = {};
        try { body = JSON.parse(opts.body || '{}'); } catch (e) { body = {}; }
        return _ensureToken().then(function () { return _sauver(rid, body); });
      }
      if (method === 'GET') return _ensureToken().then(function () { return _ouvrir(rid); });
    }
    m = url.match(/^\/api\/releves\/([^/]+)\/fichiers$/);
    if (m && method === 'GET') {
      var rid2 = decodeURIComponent(m[1]);
      return _ensureToken().then(function () { return _fichiersLister(rid2); });
    }
    m = url.match(/^\/api\/releves\/([^/]+)\/fichiers\/([^/]+)$/);
    if (m && method === 'PUT') {
      var rid3 = decodeURIComponent(m[1]);
      var nom = decodeURIComponent(m[2]);
      var body2 = {};
      try { body2 = JSON.parse(opts.body || '{}'); } catch (e) { body2 = {}; }
      return _ensureToken().then(function () { return _fichierSauver(rid3, nom, body2); });
    }
    m = url.match(/^\/api\/releves\/([^/]+)\/statut$/);
    if (m && method === 'PUT') {
      var rid4 = decodeURIComponent(m[1]);
      var body3 = {};
      try { body3 = JSON.parse(opts.body || '{}'); } catch (e) { body3 = {}; }
      return _ensureToken().then(function () { return _changerStatut(rid4, body3.statut); });
    }
    return Promise.reject(_err(404, 'Endpoint inconnu: ' + url));
  }

  /* --------------------------- SURCHARGES biblio* -------------------------- */

  function biblioInit() {
    // Le code d'accès disparaît ; on gère MSAL
    try { localStorage.removeItem('snaf_biblio_code'); } catch (e) {}
    _handleRedirect();
  }
  function biblioEnterCode() { msalLogin(); }

  function _msgErreur(e) {
    // Les erreurs MSAL/Graph utilisent .message/.errorMessage ; nos _err utilisent .msg
    if (!e) return 'Erreur inconnue';
    return String(e.msg || e.errorMessage || e.message || ('Erreur ' + (e.status || ''))).trim() || 'Erreur inconnue';
  }

  function biblioCharger() {
    biblioAfficher('login');
    var st = biblioEl('biblioStatus');
    if (st) st.textContent = 'Connexion…';
    return _ensureToken().then(_me).then(function (me) {
      biblio.user = biblio.user || { initials: me.initials, role: me.role, canDelete: me.canDelete };
      biblio.on = true;
      return biblioApi('/api/releves');
    }).then(function (data) {
      biblio.liste = data.releves || [];
      try { biblioListeCacheSauver(biblio.liste); } catch (e) {}
      biblioAfficher('lib');
      biblioAfficherListe();
    }).catch(function (err) {
      var locale = [];
      try { locale = biblioListeHorsLigne(); } catch (e) {}
      if (locale && locale.length && !(err && err.status === 403)) {
        biblio.on = false;
        if (!biblio.user) biblio.user = { initials: '?', role: 'NR', canDelete: false };
        biblio.liste = locale;
        biblioToast('Hors-ligne — liste locale (' + locale.length + ' dossier(s))');
        biblioAfficher('lib');
        biblioAfficherListe();
        return;
      }
      // Afficher la VRAIE cause (403/consentement/compte) au lieu d'un faux « hors-ligne »
      var cause = _msgErreur(err);
      biblioToast(cause);
      try { console.error('[SNAF] Échec chargement:', err); } catch (e) {}
      afficherLogin();
    });
  }

  function biblioConsulterDocument(nom) {
    if (!biblio.id) return;
    _fichierBlob(biblio.id, nom).then(function (blob) {
      var obj = URL.createObjectURL(blob);
      if (typeof window.open === 'function') window.open(obj, '_blank');
      setTimeout(function () { try { URL.revokeObjectURL(obj); } catch (e) {} }, 30000);
    }).catch(function (e) {
      if (typeof biblioToast === 'function') biblioToast((e && e.msg) || (e && e.message) || 'Impossible d\'ouvrir le document');
    });
  }

  function biblioAfficherPhotos() {
    var grid = document.getElementById('photoGrid');
    if (!grid) return;
    grid.innerHTML = '';
    var photos = [];
    try { photos = biblioFichiersPhotos(); } catch (e) {}
    if (!photos.length) {
      var v = document.createElement('div');
      v.style.cssText = 'grid-column:1 / span 2;color:#9fb6dd;text-align:center;padding:40px 0;font-size:15px';
      v.textContent = 'Aucune photo. Tapez ➕ Ajouter pour prendre ou choisir des photos.';
      grid.appendChild(v);
      return;
    }
    photos.forEach(function (f) {
      var cell = document.createElement('div');
      cell.style.cssText = 'background:#1a2438;border-radius:10px;overflow:hidden;cursor:pointer;min-height:140px;display:flex;flex-direction:column';
      var img = document.createElement('img');
      img.style.cssText = 'width:100%;height:130px;object-fit:cover;background:#2a3350';
      _fichierBlob(biblio.id, f.nom).then(function (blob) {
        img.src = URL.createObjectURL(blob);
      }).catch(function () {});
      cell.appendChild(img);
      cell.onclick = (function (n) { return function () { biblioConsulterDocument(n); }; })(f.nom || '');
      grid.appendChild(cell);
    });
  }

  function biblioUploadPdf(doc, nom) {
    if (!biblio.id || !biblio.open) return;
    var blob = null;
    try { blob = doc.output('blob'); } catch (e1) {}
    if (!blob) return;
    var reader = new FileReader();
    reader.onload = function () {
      var base64 = (reader.result || '').toString().split(',')[1];
      if (!base64) return;
      var safe = (nom || 'document.pdf').replace(/[^a-zA-Z0-9._-]/g, '_').replace(/^[^a-zA-Z0-9]+/, 'x');
      biblioApi('/api/releves/' + biblio.id + '/fichiers/' + safe, { method: 'PUT', body: JSON.stringify({ dataBase64: base64 }) })
        .then(function () {
          biblioToast('📄 PDF ajouté au dossier');
          biblioChargerFichiers();
        })
        .catch(function (err) {
          biblioToast((err && err.msg) || 'Échec ajout du PDF');
        });
    };
    try { reader.readAsDataURL(blob); } catch (e) {}
  }

  function biblioSurveillerReseau() {
    var connecte = function () { return !!(biblio.user && _token); };
    var on = function () {
      if (biblio.on === false && connecte()) {
        biblioToast('Réseau rétabli');
        try { biblioChargerFichiers(); } catch (e) {}
      }
      biblio.on = true;
      try { if (typeof biblioSync === 'function') biblioSync(); } catch (e) {}
    };
    var off = function () { biblioToast('Réseau coupé — mode hors-ligne'); biblio.on = false; };
    if (navigator.onLine === false) biblio.on = false;
    window.addEventListener('online', on);
    window.addEventListener('offline', off);
    setInterval(function () {
      if (navigator.onLine && connecte()) {
        try { if (!biblio.open && typeof biblioSync === 'function') biblioSync(); } catch (e) {}
      }
    }, 8000);
  }

  /* ---------------------------- DÉMARRAGE ---------------------------------- */

  window.biblioApi = biblioApi;
  window.biblioInit = biblioInit;
  window.biblioEnterCode = biblioEnterCode;
  window.biblioCharger = biblioCharger;
  window.biblioDeconnecter = deconnecter;
  window.biblioConsulterDocument = biblioConsulterDocument;
  window.biblioAfficherPhotos = biblioAfficherPhotos;
  window.biblioUploadPdf = biblioUploadPdf;
  window.biblioSurveillerReseau = biblioSurveillerReseau;
  window.msalLogin = msalLogin;
  window.snafGraphDebug = function () {
    return {
      account: _account ? _account.username : null,
      token: !!_token,
      cache: { dossiers: Object.keys(_cache.byId).length, ts: _cache.ts }
    };
  };

  function _start() {
    var inst = _msalInit();
    if (inst && typeof biblioAfficher === 'function') {
      _handleRedirect();
    } else {
      afficherLogin();
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', function () { setTimeout(_start, 60); });
  } else {
    setTimeout(_start, 60);
  }
})();
