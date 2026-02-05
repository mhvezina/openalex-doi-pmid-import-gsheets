// OpenAlex_DOI_PMID_Importation pour Google Sheets
// (OpenAlex_DOI_PMID_Import for Google Sheets)
// Version : 1.0.0
// Date : 2025-12-15
//
// (c) Marie-Helene Vezina, Direction des Bibliotheques, Universite de Montreal
// Distribue sous licence GPL (GNU General Public License), version 3
// Distributed under the GPL (GNU General Public License), version 3
//
// FRANCAIS:
//
// Ce script :
// 1) lit des DOI et des PMID dans l'onglet DOI/PMID (ou equivalent) d'une feuille Google Sheets liee,
// 2) normalise les identifiants (DOI en forme canonique; PMID accepte aussi une URL du type https://pubmed.ncbi.nlm.nih.gov/nnnnnn/),
// 3) les decoupe en blocs de 50 (limite documentee pour le separateur | (=OR) dans un meme filtre),
// 4) construit des requetes https://api.openalex.org/works?filter=... avec d'eventuels filtres publication_year, type, open_access.is_oa, include_xpac, etc. a partir de la feuille Parametres/Parameters,
// 5) appelle l'API d'OpenAlex via UrlFetchApp.fetch,
// 6) ecrit dans l'onglet Resultats/Results un tableau avec quelques champs (id, DOI, PMID, titre, auteurs, annee, type) ainsi que des informations sur le libre acces basees sur best_oa_location (nom et type de la source, licence, version, URL OA) et sur certains indicateurs open_access (is_oa, oa_status).
//
// Dans OpenAlex, best_oa_location represente le meilleur point d'acces libre a l'article (en general l'URL la plus stable et utile pour un texte integral en libre acces, quand disponible).
//
// L'ordre des lignes dans la feuille Resultats/Results reflete l'ordre renvoye par l'API OpenAlex, qui peut differer de l'ordre initial des identifiants (DOI/PMID) dans la feuille d'entree, et le script ne reordonne pas ces resultats.
//
// ENGLISH
//
// This script:
// 1) reads DOIs and PMIDs from the DOI/PMID (or equivalent) sheet of a linked Google Sheets file,
// 2) normalizes identifiers (DOIs into canonical form; PMIDs may also be provided as URLs such as https://pubmed.ncbi.nlm.nih.gov/nnnnnn/),
// 3) splits them into blocks of 50 (documented limit for the | (=OR) separator in a single filter),
// 4) builds https://api.openalex.org/works?filter=... queries with optional filters publication_year, type, open_access.is_oa, include_xpac, etc. from the Parametres/Parameters sheet,
// 5) calls the OpenAlex API via UrlFetchApp.fetch,
// 6) writes into the Resultats/Results sheet a table with selected fields (id, DOI, PMID, title, authors, year, type) and open access information derived from best_oa_location (source name and type, license, version, OA URL) together with key open_access indicators (is_oa, oa_status).
//
// In OpenAlex, best_oa_location represents the best open access entry point for the article (generally the most stable and useful URL to an open full text, when available).
//
// The row order in the Results sheet follows the order returned by the OpenAlex API, which may differ from the original identifier order (DOI/PMID) in the input sheet, and the script does not reorder these results.
// -------------------------------------------------------------------------------------------------------------------------------------------------------------------------

// === Menu personnalisé dans Google Sheets ===
// === Custom menu in Google Sheets ===
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🔓🔑 OpenAlex')
    .addItem('>> Mettre à jour / Update (OpenAlex)', 'runOpenAlexImport')
    .addToUi();
}

// Point d'entrée principal
// Main entry point
function runOpenAlexImport() {
  var ss = SpreadsheetApp.getActive();
  var sheetDois = ss.getSheetByName('DOI ou/or PMID');                // onglet où sont saisis les identifiants à interroger (DOI ou PMID)
  var sheetParams = ss.getSheetByName('Paramètres/Parameters');       // onglet contenant les filtres (mailto, année, type, only_oa...)
  var sheetResults = ss.getSheetByName('Résultats/Results');          // onglet où les résultats sont remplis par le script
  // sheet where DOIs to be queried are entered
  // sheet containing filters (mailto, year, type, only_oa...)
  // sheet where results are populated by the script

  if (!sheetDois || !sheetParams || !sheetResults) {
    SpreadsheetApp.getUi().alert(
      'Erreur : Les feuilles DOI, Paramètres ou Résultats sont manquantes.  / Error : DOI, Parameters and/or Results sheets are missing'
    );
    // Message d'erreur si une ou plusieurs feuilles attendues sont manquantes
    // Error message shown when one or more expected sheets are missing
    return;
  }

  var inputs = getIdInputsFromSheet_(sheetDois);
  if (inputs.length === 0) {
    SpreadsheetApp.getUi().alert('Erreur : Aucun identifiant (DOI ou PubMed ID) détecté dans la feuille DOI. / Error: No identifier (DOI or PubMed ID) found in the DOI sheet');
    // Message d'erreur si aucun identifiant n'est trouvé dans la feuille DOI
    // Error message shown when no identifier is found in the DOI sheet
    return;
  }

  var params = getParams_(sheetParams);
  var rows = fetchWorksForIds_(inputs, params);

  // Écriture dans la feuille Résultats
  // Writing results to the Results sheet
  sheetResults.clearContents();

  // Remettre en blanc tout le canevas de résultats (fond des lignes)
  // Reset the background of the results area to white
  var maxRows = sheetResults.getMaxRows();
  var maxCols = sheetResults.getMaxColumns();
  if (maxRows > 1 && maxCols > 0) {
    sheetResults
      .getRange(2, 1, maxRows - 1, maxCols)
      .setBackground('#ffffff');
  }

  // Colonnes des résultats (noms originaux d'OpenAlex + champs OA supplémentaires)
  // Result columns (original OpenAlex field names + additional OA fields)
  var header = [
    'doi',
    'pmid',
    'openalex_id',
    'title',
    'authors_display_name',
    'publication_year',
    'publication_date',
    'type',
    'is_oa',
    'oa_status',
    'source_display_name',
    'source_type',
    'best_oa_license',
    'best_oa_version',
    'oa_url'
  ];

  sheetResults.getRange(1, 1, 1, header.length).setValues([header]);

  if (rows.length > 0) {
    sheetResults
      .getRange(2, 1, rows.length, header.length)
      .setValues(rows);
  }

  if (rows.length > 0) {
    sheetResults
      .getRange(2, 1, rows.length, header.length)
      .setValues(rows);
  }

  // Appliquer un fond vert clair aux lignes avec is_oa = TRUE
  // Apply a light green background to rows where is_oa = TRUE
  var lastRow = sheetResults.getLastRow();
  if (lastRow > 1) {
    var dataRange = sheetResults.getRange(2, 1, lastRow - 1, header.length);
    var values = dataRange.getValues();
    var backgrounds = dataRange.getBackgrounds();

    for (var i = 0; i < values.length; i++) {
      var isOa = values[i][8]; // colonne I (index 8 en base 0) = is_oa
      // column I (0-based index 8) = is_oa

      var color;
      if (isOa === true || isOa === 'TRUE' || isOa === 'VRAI' || isOa === 'vrai') {
        color = '#e6fff2'; // vert clair // light green
      } else {
        color = '#ffe6e6'; // rouge clair // light red 
      }

      for (var j = 0; j < backgrounds[0].length; j++) {
        backgrounds[i][j] = color;
      }
    }

    dataRange.setBackgrounds(backgrounds);
  }
}


// === Lecture des identifiants (DOI ou PMID) ===
// === Reading identifiers (DOI or PMID) from the sheet ===
function getIdInputsFromSheet_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); // col A
  // Récupérer les valeurs à partir de la colonne A, ligne 2
  // Retrieve values from column A starting at row 2
  var inputs = [];

  values.forEach(function (row) {
    var raw = (row[0] || '').toString().trim();
    if (!raw) {
      return;
    }

    // Normaliser un éventuel DOI donné sous forme d'URL (https://doi.org/ ou https://dx.doi.org/)
    // Normalize a possible DOI given as a URL (https://doi.org/ or https://dx.doi.org/)
    var cleaned = raw.replace(/^https?:\/\/(dx\.)?doi\.org\//i, '');

    // Normaliser un éventuel PMID donné sous forme d'URL PubMed (https://pubmed.ncbi.nlm.nih.gov/nnnnnn)
    // Normalize a possible PMID given as a PubMed URL (https://pubmed.ncbi.nlm.nih.gov/nnnnnn)
    cleaned = cleaned.replace(/^https?:\/\/pubmed\.ncbi\.nlm\.nih\.gov\/([0-9]+)\/?/i, '$1');

    var kind = '';
    var value = cleaned;

    // Heuristique simple :
    // - commence par "10." -> DOI
    // - seulement des chiffres -> PMID
    // Simple heuristic:
    // - starts with "10." -> DOI
    // - only digits -> PMID
    if (/^10\.\S+/i.test(cleaned)) {
      kind = 'doi';
    } else if (/^[0-9]+$/.test(cleaned)) {
      kind = 'pmid';
    } else {
      kind = 'unknown';
    }

    inputs.push({
      raw: raw,
      kind: kind,
      value: value
    });
  });

  return inputs;
}

// === Lecture des paramètres ===
// === Reading parameters ===
function getParams_(sheetParams) {
  var data = sheetParams.getDataRange().getValues(); // supposée petite feuille de paramètres
  // Lecture de toutes les lignes de paramètres (taille attendue réduite)
  // Read all parameter rows (expected to be small)
  var map = {};

  for (var i = 1; i < data.length; i++) {
    var key = (data[i][0] || '').toString().trim();
    var val = data[i][1];
    if (key) {
      map[key] = val;
    }
  }

  return {
    mailto: map['mailto'] ? map['mailto'].toString().trim() : '',
    yearMin: map['year_min'] ? parseInt(map['year_min'], 10) : null,
    yearMax: map['year_max'] ? parseInt(map['year_max'], 10) : null,
    type: map['type'] ? map['type'].toString().trim() : '',
    onlyOa: map['only_oa'] === true || map['only_oa'] === 'TRUE'  || map['only_oa'] === 'VRAI'  || map['only_oa'] === 'vrai',
    // Interprète only_oa comme vrai pour TRUE, VRAI, vrai ou valeur booléenne true
    // Interpret only_oa as true for TRUE, VRAI, vrai or boolean true
    includeXpac: map['include_xpac'] === true || map['include_xpac'] === 'TRUE' || map['include_xpac'] === 'VRAI' || map['include_xpac'] === 'vrai'
    // Interprète include_xpac comme vrai pour TRUE, VRAI, vrai ou valeur booléenne true
    // Interpret include_xpac as true for TRUE, VRAI, vrai or boolean true
  };
}

// === Découper un tableau en blocs de taille max ===
// === Split an array into chunks of a given maximum size ===
function chunkArray_(arr, size) {
  var chunks = [];
  for (var i = 0; i < arr.length; i += size) {
    chunks.push(arr.slice(i, i + size));
  }
  return chunks;
}

// === Construction des lignes de résultat à partir d'un objet Work ===
// === Build result rows from a Work object ===
function appendWorkRow_(work, allRows) {
  // DOI normalisé (si disponible)
  // Normalized DOI (if available)
  var doiValue = '';
  if (work.doi) {
    doiValue = normalizeDoi_(work.doi);
  } else if (work.ids && work.ids.doi) {
    doiValue = normalizeDoi_(work.ids.doi);
  }

  // Transformer en URL cliquable si un DOI est présent
  // Turn into a clickable URL if a DOI is present
  if (doiValue) {
    doiValue = 'https://doi.org/' + doiValue;
  }

  // PMID (si disponible)
  // PMID (if available)
  var pmidValue = '';
  if (work.ids && work.ids.pmid) {
    pmidValue = work.ids.pmid.toString();
  }

  // OA
  // Open access information
  var oa = work.open_access || {};
  var isOaVal = oa.is_oa === true;
  var oaStatus = oa.oa_status || '';
  var oaUrl = oa.oa_url || '';

  // Date de publication
  // Publication date
  var pubDate = work.publication_date || '';

  // Nom et type de la meilleure source OA (best_oa_location)
  // Name and type of the best OA source (best_oa_location)
  var sourceName = '';
  var sourceType = '';
  var bestOaLicense = '';
  var bestOaVersion = '';
  if (work.best_oa_location) {
    if (
      work.best_oa_location.source &&
      work.best_oa_location.source.display_name
    ) {
      sourceName = work.best_oa_location.source.display_name;
    }
    if (
      work.best_oa_location.source &&
      work.best_oa_location.source.type
    ) {
      sourceType = work.best_oa_location.source.type;
    }
    bestOaLicense = work.best_oa_location.license || '';
    bestOaVersion = work.best_oa_location.version || '';
  }

  // Noms d'auteurs (limiter aux 3 premiers, puis "et al." s'il y en a plus)
  // Author names (limit to first 3, then add "et al." if there are more)
  var authorsDisplay = '';
  if (work.authorships && Array.isArray(work.authorships)) {
    var names = work.authorships
      .map(function (a) {
        if (a.author && a.author.display_name) {
          return a.author.display_name;
        }
        return '';
      })
      .filter(function (n) {
        return n !== '';
      });

    if (names.length <= 3) {
      authorsDisplay = names.join('; ');
    } else {
      var firstThree = names.slice(0, 3);
      authorsDisplay = firstThree.join('; ') + '; et al.';
    }
  }

  // Préparer l'année de publication comme texte pour éviter l'interprétation en date par Google Sheets
  // Prepare publication year as text to avoid Google Sheets interpreting it as a date
  var pubYearText = '';
  if (work.publication_year) {
    pubYearText = work.publication_year.toString();
  }

  var row = [
    doiValue,
    pmidValue,
    work.id || '',
    work.display_name || '',
    authorsDisplay,
    pubYearText,
    pubDate,
    work.type || '',
    isOaVal,
    oaStatus,
    sourceName,
    sourceType,
    bestOaLicense,
    bestOaVersion,
    oaUrl
  ];
  allRows.push(row);
}

// === Appels à l'API OpenAlex (DOI ou PMID) ===
// === Calls to the OpenAlex API (DOI or PMID) ===
function fetchWorksForIds_(inputs, params) {
  var baseUrl = 'https://api.openalex.org/works';
  var allRows = [];

  // Séparer les entrées en DOI et PMID
  // Split inputs into DOI and PMID
  var doiValues = [];
  var pmidValues = [];

  inputs.forEach(function (item) {
    if (item.kind === 'doi') {
      doiValues.push(item.value);
    } else if (item.kind === 'pmid') {
      pmidValues.push(item.value);
    } else {
      Logger.log('Identifiant non reconnu (ni DOI ni PMID) : ' + item.raw);
      // Unrecognized identifier, log and ignore
    }
  });

  // 1) Requêtes pour les DOI avec filtre doi:
  // 1) Requests for DOIs with doi: filter
  if (doiValues.length > 0) {
    var doiChunks = chunkArray_(doiValues, 50); // jusqu'à 50 DOI par requête
    // up to 50 DOIs per request

    for (var c = 0; c < doiChunks.length; c++) {
      var chunk = doiChunks[c];

      var filterParts = [];
      var doiExpr = 'doi:' + chunk.join('|');
      filterParts.push(doiExpr);

      // 2) filtre sur l'année de publication (optionnel)
      // 2) filter on publication year (optional)
      if (params.yearMin || params.yearMax) {
        var yearValue;
        if (params.yearMin && params.yearMax) {
          yearValue = params.yearMin + '-' + params.yearMax; // ex 2018-2024
        } else if (params.yearMin) {
          yearValue = '>' + params.yearMin;
        } else {
          yearValue = '<' + params.yearMax;
        }
        filterParts.push('publication_year:' + yearValue);
      }

      // 3) filtre sur le type (optionnel)
      // 3) filter on type (optional)
      if (params.type) {
        filterParts.push('type:' + params.type);
      }

      // 4) filtre sur libre accès oui ou non (optionnel)
      // 4) filter on open access yes/no (optional)
      if (params.onlyOa) {
        filterParts.push('open_access.is_oa:true');
      }

      var filterExpr = filterParts.join(',');

      // IMPORTANT : on laisse OpenAlex renvoyer l'objet complet (pas de select)
      // IMPORTANT: let OpenAlex return the full work object (no select parameter)
      var queryParts = [
        'filter=' + encodeURIComponent(filterExpr),
        'per-page=50'
      ];

      if (params.includeXpac) {
        queryParts.push('include_xpac=true');
        // Ajoute include_xpac=true à la requête si activé dans les paramètres
        // Add include_xpac=true to the request if enabled in parameters
      }

      if (params.mailto) {
        queryParts.push('mailto=' + encodeURIComponent(params.mailto));
      }

      var url = baseUrl + '?' + queryParts.join('&');

      Logger.log('Requete OpenAlex (DOI) bloc ' + (c + 1) + '/' + doiChunks.length + ': ' + url);
      // Journaliser l'URL de requête OpenAlex pour ce bloc (DOI)
      // Log the OpenAlex request URL for this DOI block

      var response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true
      });

      var code = response.getResponseCode();
      Logger.log('HTTP ' + code + ' pour ce bloc (DOI).');
      // Journaliser le code HTTP pour ce bloc
      // Log HTTP status code for this DOI block

      if (code !== 200) {
        Logger.log('Erreur : contenu de la réponse (DOI) : ' + response.getContentText());
        // Journaliser le contenu de la réponse en cas d'erreur (code HTTP != 200)
        // Log error body when status is not 200
        continue;
      }

      var text = response.getContentText();
      var json = JSON.parse(text);

      if (json.meta && typeof json.meta.count !== 'undefined') {
        Logger.log('meta.count (DOI) = ' + json.meta.count);
      }

      if (!json.results || !Array.isArray(json.results)) {
        Logger.log('Erreur : Pas de tableau résultats dans la réponse (DOI).');
        continue;
      }

      json.results.forEach(function (work) {
        appendWorkRow_(work, allRows);
      });

      // Petite pause pour rester sous les limites de l'API
      // Small delay to stay well under API rate limits
      Utilities.sleep(200);
    }
  }

  // 2) Requêtes pour les PMID avec filtre ids.pmid:
  // 2) Requests for PMIDs with ids.pmid: filter
  if (pmidValues.length > 0) {
    var pmidChunks = chunkArray_(pmidValues, 50); // jusqu'à 50 PMID par requête
    // up to 50 PMIDs per request

    for (var k = 0; k < pmidChunks.length; k++) {
      var pmidChunk = pmidChunks[k];

      var filterPartsPmid = [];
      var pmidExpr = 'ids.pmid:' + pmidChunk.join('|');
      filterPartsPmid.push(pmidExpr);

      if (params.yearMin || params.yearMax) {
        var yearValueP;
        if (params.yearMin && params.yearMax) {
          yearValueP = params.yearMin + '-' + params.yearMax;
        } else if (params.yearMin) {
          yearValueP = '>' + params.yearMin;
        } else {
          yearValueP = '<' + params.yearMax;
        }
        filterPartsPmid.push('publication_year:' + yearValueP);
      }

      if (params.type) {
        filterPartsPmid.push('type:' + params.type);
      }

      if (params.onlyOa) {
        filterPartsPmid.push('open_access.is_oa:true');
      }

      var filterExprPmid = filterPartsPmid.join(',');

      var queryPartsPmid = [
        'filter=' + encodeURIComponent(filterExprPmid),
        'per-page=50'
      ];

      if (params.includeXpac) {
        queryPartsPmid.push('include_xpac=true');
      }

      if (params.mailto) {
        queryPartsPmid.push('mailto=' + encodeURIComponent(params.mailto));
      }

      var urlPmid = baseUrl + '?' + queryPartsPmid.join('&');

      Logger.log('Requete OpenAlex (PMID) bloc ' + (k + 1) + '/' + pmidChunks.length + ': ' + urlPmid);
      // Journaliser l'URL de requête OpenAlex pour ce bloc (PMID)
      // Log the OpenAlex request URL for this PMID block

      var responsePmid = UrlFetchApp.fetch(urlPmid, {
        muteHttpExceptions: true
      });

      var codePmid = responsePmid.getResponseCode();
      Logger.log('HTTP ' + codePmid + ' pour ce bloc (PMID).');

      if (codePmid !== 200) {
        Logger.log('Erreur : contenu de la réponse (PMID) : ' + responsePmid.getContentText());
        continue;
      }

      var textPmid = responsePmid.getContentText();
      var jsonPmid = JSON.parse(textPmid);

      if (jsonPmid.meta && typeof jsonPmid.meta.count !== 'undefined') {
        Logger.log('meta.count (PMID) = ' + jsonPmid.meta.count);
      }

      if (!jsonPmid.results || !Array.isArray(jsonPmid.results)) {
        Logger.log('Erreur : Pas de tableau résultats dans la réponse (PMID).');
        continue;
      }

      jsonPmid.results.forEach(function (work) {
        appendWorkRow_(work, allRows);
      });

      Utilities.sleep(200);
    }
  }

  return allRows;
}


// Normalisation d'un DOI (retirer https://doi.org/ si présent)
// DOI normalization (remove https://doi.org/ prefix if present)
function normalizeDoi_(value) {
  var v = (value || '').toString().trim();
  return v.replace(/^https?:\/\/(dx\.)?doi\.org\//i, '');
}

