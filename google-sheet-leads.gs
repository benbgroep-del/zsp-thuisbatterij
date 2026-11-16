/**
 * ZSP thuisbatterij – aanvragen van de landingspagina naar Google Sheet + e-mail + Meta Conversions API.
 *
 * De pagina stuurt drie soorten berichten, elk met hetzelfde lead_id:
 *   lead        het formulier is verstuurd (naam, telefoon, adres, herkomst)  -> nieuwe rij + e-mail + Meta 'Lead'
 *   verrijking  de kwalificatievragen (tijdlijn, omvormer, verbruik, ...)     -> dezelfde rij aanvullen (+ Meta 'Kwalificatie' als klaar)
 *   afspraak    de klant heeft een moment in Calendly gekozen                 -> status "Afspraak gepland" + Meta 'Schedule'
 * Alles komt ook in het tabblad "Log" (het volledige bericht), zodat er nooit iets verloren gaat.
 *
 * Installeren (eenmalig, ± 5 minuten):
 *   1. Open de Google Sheet "ZSP leads thuisbatterij" → Extensies → Apps Script.
 *   2. Vervang de code door dit bestand en klik op Opslaan.
 *   3. Vul hieronder MELDING_AAN in (of laat leeg: dan gaat de melding naar uw eigen account).
 *   4. Implementeren → Nieuwe implementatie → type "Web-app"
 *        Uitvoeren als: Ik      Wie heeft toegang: Iedereen
 *      → Implementeren → toegang verlenen.
 *   5. Kopieer de URL van de web-app (eindigt op /exec) en geef die door: die komt in config.json bij leadEndpoint.
 *   Test: open die URL in de browser; u ziet "ZSP lead-ontvanger werkt".
 *   Na een wijziging in deze code: Implementeren → Implementaties beheren → potlood → Versie: "Nieuwe versie"
 *   → Implementeren (de URL blijft gelijk).
 *
 * Meta Conversions API (server) – eenmalig instellen:
 *   a. Events Manager → Pixel ZSP → Instellingen → Conversions API → "Toegangstoken genereren". Kopieer het token.
 *   b. Apps Script → Projectinstellingen (tandwiel) → Scripteigenschappen → Scripteigenschap toevoegen:
 *        META_CAPI_TOKEN   het token uit stap a
 *        META_TEST_CODE    (alleen tijdens testen) de code uit Events Manager → Gebeurtenissen testen, bv. TEST12345.
 *                          Haal deze weg na de test, anders tellen echte leads als test.
 *   c. Kies bovenin de functie testMetaVerbinding en klik op Uitvoeren. Geef toestemming voor "verbinding met een
 *      externe service". Het logboek moet melden: "Token werkt voor dataset Pixel ZSP".
 *   d. Implementeer een nieuwe versie (zie hierboven).
 *   Zonder token stuurt dit script niets naar Meta; de leads komen dan gewoon in de Sheet.
 *
 * Toestemming: het script stuurt alleen naar Meta als de bezoeker daar toestemming voor gaf:
 *   'cookies'    marketingcookies geaccepteerd in de cookiebanner, of
 *   'formulier'  het (optionele, niet vooraf aangevinkte) vinkje voor advertentiemeting in het formulier.
 * Zonder toestemming ('nee') gaat er niets naar Meta, ook niet vanaf de server.
 */
var MELDING_AAN = '';            // bijv. 'verkoop@zsp.nl' of meerdere gescheiden door een komma; leeg = uw eigen account
var TABBLAD_LEADS = 'Leads';
var TABBLAD_LOG = 'Log';

var META_PIXEL_ID = '1872449343280965';   // dataset "Pixel ZSP"
var META_API_VERSIE = 'v25.0';            // geldig tot ten minste juli 2028

var KOLOMMEN = [
  'ontvangen', 'lead_id', 'status', 'voornaam', 'achternaam', 'telefoon', 'email',
  'postcode', 'huisnummer', 'toevoeging', 'straat', 'woonplaats', 'adres_check', 'zonnepanelen',
  'tijdlijn', 'omvormer', 'aantal_panelen', 'aansluiting', 'stroomverbruik', 'gasverbruik', 'apparaten', 'interesse',
  'kwalificatie_klaar', 'afspraak_gepland_op', 'calendly_event', 'calendly_genodigde',
  'pagina', 'persona', 'concept_id', 'utm_source', 'utm_medium', 'utm_campaign', 'utm_content', 'utm_term',
  'campaign_id', 'adset_id', 'ad_id', 'fbclid', 'gclid', 'toestemming_marketing', 'pagina_url',
  // sinds 16-11-2026: Conversions API vanaf de server
  'toestemming_meta', 'meta_fbc', 'meta_fbp', 'user_agent', 'meta_capi'
];

function doGet() {
  return ContentService.createTextOutput('ZSP lead-ontvanger werkt.');
}

function doPost(e) {
  var taken = [];                                        // berichten voor Meta: pas na het vrijgeven van de lock versturen
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  var ss;
  try {
    var data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!data || !data.lead_id || ['lead', 'verrijking', 'afspraak'].indexOf(data.type) < 0) return antwoord(false, 'ongeldig bericht');
    ss = SpreadsheetApp.getActiveSpreadsheet();
    schrijfLog(ss, data);
    if (data.type === 'lead') {
      if (nieuweLead(ss, data)) taken.push({ naam: 'Lead', id: data.lead_id, waarde: (data.waarde || {}).lead, valuta: (data.waarde || {}).valuta });
    }
    if (data.type === 'verrijking') {
      werkBij(ss, data.lead_id, verrijkingVelden(data));
      if (data.klaar) taken.push({ naam: 'Kwalificatie', id: data.lead_id + '-kwalificatie', waarde: data.waarde, valuta: data.valuta,
        extra: { tijdlijn: (data.verrijking || {}).tijdlijn || '' } });
    }
    if (data.type === 'afspraak') {
      var c = data.calendly || {};
      werkBij(ss, data.lead_id, { status: 'Afspraak gepland', afspraak_gepland_op: data.gepland_at || new Date().toISOString(),
        calendly_event: c.event_uri || '', calendly_genodigde: c.invitee_uri || '' });
      taken.push({ naam: 'Schedule', id: data.lead_id + '-afspraak', waarde: data.waarde, valuta: data.valuta });
    }
  } catch (err) {
    return antwoord(false, String(err));
  } finally {
    lock.releaseLock();
  }
  /* Meta mag de aanvraag nooit laten mislukken: elke fout komt alleen in de kolom meta_capi. */
  taken.forEach(function (t) { try { metaVoorLead(ss, data.lead_id, t); } catch (err) { console.error('Meta: ' + err); } });
  return antwoord(true);
}

/* Geeft true terug als er een nieuwe rij is gemaakt (false bij een dubbel bericht). */
function nieuweLead(ss, d) {
  if (vindRij(blad(ss), d.lead_id)) return false;        // dubbel verstuurd: niets doen
  var a = d.adres || {}, k = d.contact || {}, u = d.utm || {}, ids = d.klik_ids || {}, t = d.toestemming || {}, m = d.meta || {};
  var metaOk = metaToestemming(t);
  var v = {
    ontvangen: new Date(), lead_id: d.lead_id, status: 'Nieuw',
    voornaam: k.voornaam, achternaam: k.achternaam, telefoon: k.telefoon, email: k.email,
    postcode: a.postcode, huisnummer: a.huisnummer, toevoeging: a.toevoeging, straat: a.straat, woonplaats: a.woonplaats,
    adres_check: a.adres_check, zonnepanelen: d.zonnepanelen,
    pagina: d.page_variant, persona: d.persona, concept_id: d.concept_id,
    utm_source: u.utm_source, utm_medium: u.utm_medium, utm_campaign: u.utm_campaign, utm_content: u.utm_content, utm_term: u.utm_term,
    campaign_id: u.campaign_id, adset_id: u.adset_id, ad_id: u.ad_id, fbclid: ids.fbclid, gclid: ids.gclid,
    toestemming_marketing: t.cookies_marketing ? 'ja' : 'nee', pagina_url: d.page_url,
    toestemming_meta: metaOk,
    meta_fbc: metaOk !== 'nee' ? m.fbc : '', meta_fbp: metaOk !== 'nee' ? m.fbp : '',
    user_agent: metaOk !== 'nee' ? d.user_agent : '',
    meta_capi: metaOk === 'nee' ? 'niet verstuurd (geen toestemming)' : ''
  };
  blad(ss).appendRow(KOLOMMEN.map(function (kol) { return veilig(v[kol]); }));
  meld(ss, v);
  return true;
}

/* 'cookies', 'formulier' of 'nee'. Oudere pagina's sturen alleen cookies_marketing mee. */
function metaToestemming(t) {
  t = t || {};
  if (t.meta_meting === 'cookies' || t.meta_meting === 'formulier') return t.meta_meting;
  return t.cookies_marketing ? 'cookies' : 'nee';
}

function verrijkingVelden(d) {
  var v = d.verrijking || {};
  return {
    tijdlijn: v.tijdlijn, omvormer: v.omvormer, aantal_panelen: v.aantal_panelen, aansluiting: v.aansluiting,
    stroomverbruik: v.stroomverbruik, gasverbruik: v.gasverbruik,
    apparaten: (v.apparaten || []).join(', '), interesse: (v.interesse || []).join(', '),
    kwalificatie_klaar: d.klaar ? 'ja' : 'bezig'
  };
}

function werkBij(ss, leadId, velden) {
  var sh = blad(ss), rij = vindRij(sh, leadId);
  if (!rij) {                                            // bericht kwam vóór de lead binnen: lege rij met het lead_id
    sh.appendRow(KOLOMMEN.map(function (kol) { return kol === 'lead_id' ? leadId : kol === 'ontvangen' ? new Date() : ''; }));
    rij = sh.getLastRow();
  }
  Object.keys(velden).forEach(function (kol) {
    var i = KOLOMMEN.indexOf(kol);
    if (i >= 0 && velden[kol] !== undefined && velden[kol] !== null && velden[kol] !== '') sh.getRange(rij, i + 1).setValue(veilig(velden[kol]));
  });
}

function blad(ss) {
  var sh = ss.getSheetByName(TABBLAD_LEADS) || ss.insertSheet(TABBLAD_LEADS);
  if (sh.getLastRow() === 0) { sh.appendRow(KOLOMMEN); sh.setFrozenRows(1); sh.getRange(1, 1, 1, KOLOMMEN.length).setFontWeight('bold'); }
  else zorgVoorKoppen(sh);
  return sh;
}

/* Nieuwe kolommen achteraan de bestaande kopregel zetten (eenmalig na een update van dit script). */
function zorgVoorKoppen(sh) {
  var n = sh.getLastColumn();
  if (n >= KOLOMMEN.length) return;
  var kop = sh.getRange(1, 1, 1, n).getValues()[0];
  KOLOMMEN.forEach(function (k, i) {
    if (i >= n && kop.indexOf(k) < 0) sh.getRange(1, i + 1).setValue(k).setFontWeight('bold');
  });
}

function vindRij(sh, leadId) {
  if (sh.getLastRow() < 2) return 0;
  var ids = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) if (ids[i][0] === leadId) return i + 2;
  return 0;
}

function rijObject(sh, rij) {
  var w = sh.getRange(rij, 1, 1, KOLOMMEN.length).getValues()[0], o = {};
  KOLOMMEN.forEach(function (k, i) { o[k] = w[i]; });
  return o;
}

function schrijfLog(ss, d) {
  var sh = ss.getSheetByName(TABBLAD_LOG) || ss.insertSheet(TABBLAD_LOG);
  if (sh.getLastRow() === 0) sh.appendRow(['ontvangen', 'type', 'lead_id', 'bericht']);
  sh.appendRow([new Date(), d.type, veilig(d.lead_id), veilig(JSON.stringify(d).slice(0, 45000))]);
}

function meld(ss, v) {
  var aan = MELDING_AAN || Session.getEffectiveUser().getEmail();
  if (!aan) return;
  var regels = [
    'Nieuwe aanvraag thuisbatterij via de landingspagina.', '',
    'Naam: ' + [v.voornaam, v.achternaam].join(' '),
    'Telefoon: ' + (v.telefoon || ''),
    'E-mail: ' + (v.email || ''),
    'Adres: ' + [v.straat, v.huisnummer, v.toevoeging].join(' ') + ', ' + (v.postcode || '') + ' ' + (v.woonplaats || ''),
    'Zonnepanelen: ' + (v.zonnepanelen || ''),
    'Herkomst: ' + [v.utm_source, v.utm_campaign, v.utm_content].filter(String).join(' / '),
    '', 'Bel binnen 1 werkdag. Alle gegevens: ' + ss.getUrl()
  ];
  MailApp.sendEmail(aan, 'Nieuwe aanvraag thuisbatterij: ' + [v.voornaam, v.achternaam].join(' ') + ' (' + (v.postcode || '') + ')', regels.join('\n'));
}

/* ---------- Meta Conversions API ----------
   Zelfde eventnaam en event_id als de pixel in de browser, dus Meta telt elk event één keer:
   Lead = lead_id, Kwalificatie = lead_id-kwalificatie, Schedule = lead_id-afspraak. */
function metaVoorLead(ss, leadId, taak) {
  var sh = blad(ss), rij = vindRij(sh, leadId);
  if (!rij) return;
  var o = rijObject(sh, rij);
  var log = String(o.meta_capi || '');
  if (o.toestemming_meta !== 'cookies' && o.toestemming_meta !== 'formulier') {
    if (!log) schrijfMetaStatus(sh, rij, '', 'niet verstuurd (geen toestemming)');
    return;
  }
  if (new RegExp('(^| · )' + taak.naam + ' (ok|test ok)').test(log)) return;   // al eerder gelukt: niet nog eens
  var uitslag = stuurNaarMeta(metaEvent(o, taak));
  schrijfMetaStatus(sh, rij, log, taak.naam + ' ' + uitslag);
}

function metaEvent(o, taak) {
  var custom = { pagina: o.pagina || '', persona: o.persona || 'geen', concept_id: o.concept_id || '' };
  if (typeof taak.waarde === 'number' && isFinite(taak.waarde)) { custom.value = taak.waarde; custom.currency = taak.valuta || 'EUR'; }
  Object.keys(taak.extra || {}).forEach(function (k) { if (taak.extra[k]) custom[k] = taak.extra[k]; });
  return {
    event_name: taak.naam,
    event_time: Math.floor(Date.now() / 1000),
    event_id: taak.id,
    action_source: 'website',
    event_source_url: o.pagina_url || '',
    user_data: metaGebruiker(o),
    custom_data: custom
  };
}

/* Klantgegevens, genormaliseerd en gehasht zoals de pagina dat ook doet (SHA-256, kleine letters). */
function metaGebruiker(o) {
  var ud = {}, zet = function (k, v) { if (v) ud[k] = [sha256(v)]; };
  zet('em', String(o.email || '').trim().toLowerCase());
  zet('ph', normTelefoon(o.telefoon));
  zet('fn', String(o.voornaam || '').trim().toLowerCase());
  zet('ln', String(o.achternaam || '').trim().toLowerCase());
  zet('zp', String(o.postcode || '').toLowerCase().replace(/\s+/g, ''));
  zet('ct', String(o.woonplaats || '').toLowerCase().replace(/[^a-z]/g, ''));
  zet('country', 'nl');
  zet('external_id', String(o.lead_id || ''));
  if (o.meta_fbc) ud.fbc = String(o.meta_fbc);
  if (o.meta_fbp) ud.fbp = String(o.meta_fbp);
  if (o.user_agent) ud.client_user_agent = String(o.user_agent);
  return ud;
}

function normTelefoon(p) {
  var d = String(p || '').replace(/\D/g, '');
  if (d.indexOf('0031') === 0) d = d.slice(2);
  else if (d.indexOf('31') === 0) d = d;
  else if (d.indexOf('0') === 0) d = '31' + d.slice(1);
  else if (d.length === 9) d = '31' + d;               // Sheets maakte van 0612345678 het getal 612345678
  return d.length >= 11 ? d : '';
}

function sha256(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

/* Geeft 'ok', 'test ok', 'geen token' of 'fout ...' terug. */
function stuurNaarMeta(ev) {
  var p = PropertiesService.getScriptProperties();
  var token = p.getProperty('META_CAPI_TOKEN');
  if (!token) return 'niet verstuurd (geen token)';
  var test = p.getProperty('META_TEST_CODE');
  var body = { data: [ev], access_token: token };
  if (test) body.test_event_code = test;
  var res = UrlFetchApp.fetch('https://graph.facebook.com/' + META_API_VERSIE + '/' + META_PIXEL_ID + '/events', {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true
  });
  var code = res.getResponseCode(), tekst = res.getContentText();
  if (code === 200) return test ? 'test ok' : 'ok';
  var fout = '';
  try { fout = (JSON.parse(tekst).error || {}).message || ''; } catch (e) { fout = tekst; }
  return 'fout ' + code + ': ' + String(fout).slice(0, 180);
}

function schrijfMetaStatus(sh, rij, oud, nieuw) {
  var kol = KOLOMMEN.indexOf('meta_capi') + 1;
  sh.getRange(rij, kol).setValue(veilig(oud ? oud + ' · ' + nieuw : nieuw));
}

/* Handmatig uitvoeren in de editor. Stuurt één testevent ('ZspVerbindingstest') met een testcode: dat komt alleen
   in "Gebeurtenissen testen" en telt niet mee in rapportage of optimalisatie. */
function testMetaVerbinding() {
  var p = PropertiesService.getScriptProperties();
  var token = p.getProperty('META_CAPI_TOKEN');
  if (!token) { console.log('Geen META_CAPI_TOKEN in de scripteigenschappen.'); return; }
  var test = p.getProperty('META_TEST_CODE');
  var body = {
    access_token: token, test_event_code: test || 'TEST_ZSP_VERBINDING',
    data: [{ event_name: 'ZspVerbindingstest', event_time: Math.floor(Date.now() / 1000), action_source: 'system_generated',
      user_data: { external_id: [sha256('zsp-verbindingstest')] } }]
  };
  var res = UrlFetchApp.fetch('https://graph.facebook.com/' + META_API_VERSIE + '/' + META_PIXEL_ID + '/events', {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true
  });
  var j = {};
  try { j = JSON.parse(res.getContentText()); } catch (e) {}
  if (res.getResponseCode() === 200) console.log('Token werkt voor dataset Pixel ZSP (' + META_PIXEL_ID + '): ' + (j.events_received || 0) + ' testevent ontvangen.');
  else console.log('Token werkt niet (' + res.getResponseCode() + '): ' + ((j.error || {}).message || res.getContentText()));
  console.log(test ? 'Testmodus AAN (META_TEST_CODE = ' + test + '): leads gaan als test naar Meta. Haal die weg na de test.'
                   : 'Testmodus uit: leads tellen echt mee in Meta.');
}

/* Tekst die met = + - @ begint, zou Sheets als formule lezen; een nummer met een voorloopnul (0612345678) zou
   Sheets als getal opslaan en de 0 kwijtraken. Beide altijd als gewone tekst opslaan. */
function veilig(x) {
  if (x === undefined || x === null) return '';
  if (x instanceof Date || typeof x === 'number' || typeof x === 'boolean') return x;
  var s = String(x).slice(0, 2000);
  return /^[=+\-@]/.test(s) || /^0\d/.test(s) ? "'" + s : s;
}

function antwoord(ok, fout) {
  return ContentService.createTextOutput(JSON.stringify(ok ? { ok: true } : { ok: false, fout: fout })).setMimeType(ContentService.MimeType.JSON);
}
