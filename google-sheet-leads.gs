/**
 * ZSP thuisbatterij – aanvragen van de landingspagina naar Google Sheet + e-mail.
 *
 * De pagina stuurt drie soorten berichten, elk met hetzelfde lead_id:
 *   lead        het formulier is verstuurd (naam, telefoon, adres, herkomst)  -> nieuwe rij + e-mail
 *   verrijking  de kwalificatievragen (tijdlijn, omvormer, verbruik, ...)     -> dezelfde rij aanvullen
 *   afspraak    de klant heeft een moment in Calendly gekozen                 -> status "Afspraak gepland"
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
 *   Na een wijziging in deze code: Implementeren → Implementaties beheren → bewerken → Nieuwe versie (de URL blijft gelijk).
 */
var MELDING_AAN = '';            // bijv. 'verkoop@zsp.nl' of meerdere gescheiden door een komma; leeg = uw eigen account
var TABBLAD_LEADS = 'Leads';
var TABBLAD_LOG = 'Log';

var KOLOMMEN = [
  'ontvangen', 'lead_id', 'status', 'voornaam', 'achternaam', 'telefoon', 'email',
  'postcode', 'huisnummer', 'toevoeging', 'straat', 'woonplaats', 'adres_check', 'zonnepanelen',
  'tijdlijn', 'omvormer', 'aantal_panelen', 'aansluiting', 'stroomverbruik', 'gasverbruik', 'apparaten', 'interesse',
  'kwalificatie_klaar', 'afspraak_gepland_op', 'calendly_event', 'calendly_genodigde',
  'pagina', 'persona', 'concept_id', 'utm_source', 'utm_medium', 'utm_campaign', 'utm_content', 'utm_term',
  'campaign_id', 'adset_id', 'ad_id', 'fbclid', 'gclid', 'toestemming_marketing', 'pagina_url'
];

function doGet() {
  return ContentService.createTextOutput('ZSP lead-ontvanger werkt.');
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!data || !data.lead_id || ['lead', 'verrijking', 'afspraak'].indexOf(data.type) < 0) return antwoord(false, 'ongeldig bericht');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    schrijfLog(ss, data);
    if (data.type === 'lead') nieuweLead(ss, data);
    if (data.type === 'verrijking') werkBij(ss, data.lead_id, verrijkingVelden(data));
    if (data.type === 'afspraak') {
      var c = data.calendly || {};
      werkBij(ss, data.lead_id, { status: 'Afspraak gepland', afspraak_gepland_op: data.gepland_at || new Date().toISOString(),
        calendly_event: c.event_uri || '', calendly_genodigde: c.invitee_uri || '' });
    }
    return antwoord(true);
  } catch (err) {
    return antwoord(false, String(err));
  } finally {
    lock.releaseLock();
  }
}

function nieuweLead(ss, d) {
  if (vindRij(blad(ss), d.lead_id)) return;               // dubbel verstuurd: niets doen
  var a = d.adres || {}, k = d.contact || {}, u = d.utm || {}, ids = d.klik_ids || {}, t = d.toestemming || {};
  var v = {
    ontvangen: new Date(), lead_id: d.lead_id, status: 'Nieuw',
    voornaam: k.voornaam, achternaam: k.achternaam, telefoon: k.telefoon, email: k.email,
    postcode: a.postcode, huisnummer: a.huisnummer, toevoeging: a.toevoeging, straat: a.straat, woonplaats: a.woonplaats,
    adres_check: a.adres_check, zonnepanelen: d.zonnepanelen,
    pagina: d.page_variant, persona: d.persona, concept_id: d.concept_id,
    utm_source: u.utm_source, utm_medium: u.utm_medium, utm_campaign: u.utm_campaign, utm_content: u.utm_content, utm_term: u.utm_term,
    campaign_id: u.campaign_id, adset_id: u.adset_id, ad_id: u.ad_id, fbclid: ids.fbclid, gclid: ids.gclid,
    toestemming_marketing: t.cookies_marketing ? 'ja' : 'nee', pagina_url: d.page_url
  };
  blad(ss).appendRow(KOLOMMEN.map(function (kol) { return veilig(v[kol]); }));
  meld(ss, v);
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
  return sh;
}

function vindRij(sh, leadId) {
  if (sh.getLastRow() < 2) return 0;
  var ids = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) if (ids[i][0] === leadId) return i + 2;
  return 0;
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

/* Tekst die met = + - @ begint, zou Sheets als formule lezen: altijd als gewone tekst opslaan. */
function veilig(x) {
  if (x === undefined || x === null) return '';
  if (x instanceof Date || typeof x === 'number' || typeof x === 'boolean') return x;
  var s = String(x).slice(0, 2000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function antwoord(ok, fout) {
  return ContentService.createTextOutput(JSON.stringify(ok ? { ok: true } : { ok: false, fout: fout })).setMimeType(ContentService.MimeType.JSON);
}
