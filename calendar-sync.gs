/**
 * לו"ז המעבר ← Google Calendar
 *
 * רץ אצל דור ב-Google Apps Script, פעם בדקה. מושך מ-Firestore פגישות שנוספו או
 * השתנו באתר, ויוצר / מעדכן / מוחק אירוע ביומן של דור, עם זימון לבר.
 *
 * התקנה (פעם אחת, מהמחשב):
 *   1. script.google.com ← "פרויקט חדש". מחקו את מה שיש ותדביקו את כל הקובץ הזה.
 *   2. שמירה (Ctrl+S). למעלה בוחרים את הפונקציה setup ולוחצים "הפעלה" (Run).
 *   3. מאשרים הרשאות: Review permissions ← בוחרים את החשבון ← Advanced ←
 *      Go to ... (unsafe) ← Allow. זה הסקריפט שלכם, האזהרה נורמלית.
 *   4. זהו. מעכשיו כל פגישה שנרשמת באתר מגיעה ליומן תוך דקה.
 */

const PROJECT = 'moving-dor-bar';
const API_KEY = 'AIzaSyC-IENp0i5ZS0NLYWSKLzkRU6pWxZFpL6M';
const GUESTS = ['barkoren1912@gmail.com'];      // דור הוא המארגן, אז האירוע נכנס ישר ליומן שלו
const TZ = 'Asia/Jerusalem';
const SITE = 'https://dorbi222.github.io/moving/#schedule';
const WHO = { dor: 'דור', bar: 'בר', both: 'שנינו' };
const FS = 'https://firestore.googleapis.com/v1/';
const DOCS = FS + 'projects/' + PROJECT + '/databases/(default)/documents';

/** מריצים פעם אחת. יוצר טריגר שרץ כל דקה ומבצע סנכרון ראשון. */
function setup() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncEvents')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('syncEvents').timeBased().everyMinutes(1).create();
  syncEvents();
  console.log('מוכן. הסנכרון ירוץ כל דקה.');
}

function syncEvents() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    const res = UrlFetchApp.fetch(DOCS + ':runQuery?key=' + API_KEY, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ structuredQuery: {
        from: [{ collectionId: 'moving_events' }],
        where: { fieldFilter: { field: { fieldPath: 'needsSync' }, op: 'EQUAL', value: { booleanValue: true } } },
      } }),
    });
    if (res.getResponseCode() !== 200) throw new Error('Firestore ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 300));
    const rows = JSON.parse(res.getContentText()).filter(r => r.document);
    if (!rows.length) return;
    const cal = CalendarApp.getDefaultCalendar();
    rows.forEach(r => {
      try { syncOne(cal, r.document); }
      catch (e) {
        console.error(e);
        patch(r.document, { syncState: 'error', syncError: String(e.message || e).slice(0, 160), needsSync: false }, false);
      }
    });
  } finally {
    lock.releaseLock();
  }
}

function syncOne(cal, doc) {
  const e = fromFields(doc.fields || {});
  let ev = e.calId ? findEvent(cal, e.calId) : null;

  if (e.deleted) {
    if (ev) ev.deleteEvent();
    UrlFetchApp.fetch(FS + doc.name + '?key=' + API_KEY, { method: 'delete', muteHttpExceptions: true });
    return;
  }
  if (!e.title || !e.date || !e.start) throw new Error('חסרים שם, תאריך או שעה');

  const start = Utilities.parseDate(e.date + ' ' + e.start, TZ, 'yyyy-MM-dd HH:mm');
  let end = Utilities.parseDate(e.date + ' ' + (e.end || e.start), TZ, 'yyyy-MM-dd HH:mm');
  if (end <= start) end = new Date(start.getTime() + 60 * 60 * 1000);
  const description = describe(e);
  const location = e.location || '';

  if (!ev) {
    ev = cal.createEvent(e.title, start, end, { description, location, guests: GUESTS.join(','), sendInvites: true });
  } else {
    ev.setTitle(e.title);
    ev.setTime(start, end);
    ev.setDescription(description);
    ev.setLocation(location);
    GUESTS.forEach(g => { if (!ev.getGuestByEmail(g)) ev.addGuest(g); });
  }

  const ok = patch(doc, { calId: ev.getId(), syncState: 'synced', syncError: '', syncedAt: Date.now(), needsSync: false }, true);
  // אם מישהו ערך את הפגישה באתר בזמן הסנכרון: שומרים רק את מזהה האירוע, והריצה הבאה תעדכן
  if (!ok) patch(doc, { calId: ev.getId() }, false);
}

function findEvent(cal, id) {
  try { return cal.getEventById(id); } catch (e) { return null; }
}

function describe(e) {
  return [
    e.contact && 'ספק / איש קשר: ' + e.contact,
    e.phone && 'טלפון: ' + e.phone,
    e.who && WHO[e.who] && 'מי פוגש: ' + WHO[e.who],
    e.notes,
    'נרשם בלו"ז המעבר: ' + SITE,
  ].filter(Boolean).join('\n');
}

/** כותב שדות למסמך. guard=true: רק אם המסמך לא השתנה מאז שקראנו אותו. */
function patch(doc, fields, guard) {
  const mask = Object.keys(fields).map(k => 'updateMask.fieldPaths=' + encodeURIComponent(k)).join('&');
  let url = FS + doc.name + '?key=' + API_KEY + '&' + mask;
  if (guard) url += '&currentDocument.updateTime=' + encodeURIComponent(doc.updateTime);
  const res = UrlFetchApp.fetch(url, {
    method: 'patch', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ fields: toFields(fields) }),
  });
  return res.getResponseCode() === 200;
}

function fromFields(f) {
  const out = {};
  Object.keys(f).forEach(k => { out[k] = fromValue(f[k]); });
  return out;
}
function fromValue(v) {
  if ('stringValue' in v) return v.stringValue;
  if ('booleanValue' in v) return v.booleanValue;
  if ('integerValue' in v) return Number(v.integerValue);
  if ('doubleValue' in v) return v.doubleValue;
  if ('nullValue' in v) return null;
  if ('mapValue' in v) return fromFields(v.mapValue.fields || {});
  if ('arrayValue' in v) return (v.arrayValue.values || []).map(fromValue);
  return null;
}
function toFields(o) {
  const out = {};
  Object.keys(o).forEach(k => {
    const v = o[k];
    if (typeof v === 'boolean') out[k] = { booleanValue: v };
    else if (typeof v === 'number') out[k] = Number.isInteger(v) ? { integerValue: String(v) } : { doubleValue: v };
    else if (v == null) out[k] = { nullValue: null };
    else out[k] = { stringValue: String(v) };
  });
  return out;
}
