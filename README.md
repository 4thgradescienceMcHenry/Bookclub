# Bookclub
4th Grade Book Club
/**
 * Curious Minds Book Club — automatic certificates and coupons
 * ------------------------------------------------------------
 * Works with the Google Form AND the GitHub Pages website (see doPost / doGet below).
 * Paste this whole file into Extensions > Apps Script in the Book Club Google Sheet,
 * save, then choose "Book Club > 1. Set up (run once)" from the sheet's menu.
 *
 * What it does:
 *  - Builds the Google Form book log (3 required written questions + rating questions).
 *  - Every time a student logs a book, checks the student ID against the Roster,
 *    blocks duplicates and too-short answers, numbers the book, and (in Auto mode)
 *    emails the Learning Coach. Every COUPON_EVERY books (default 5) it attaches a Skip-a-Quiz Pass coupon;
 *    every CERTIFICATE_EVERY books (default 20) it attaches a Reading Champion certificate.
 *    Other books get a short progress email (turn off with PROGRESS_EMAILS = No).
 *  - In Review mode, nothing is sent until you type "Yes" in the Approve column and
 *    choose "Book Club > Send approved & queued now".
 *  - If the daily email limit is reached, entries wait in the queue and send automatically later.
 */

const SHEETS = { settings: 'Settings', roster: 'Roster', log: 'Book Log', rewards: 'Rewards' };

// Book Log columns (1-based)
const COL = {
  timestamp: 1, id: 2, name: 3, cls: 4, coachEmail: 5, title: 6, author: 7, type: 8,
  q1: 9, q2: 10, q3: 11, rating: 12, recommend: 13, science: 14,
  bookNum: 15, coupon: 16, status: 17, approve: 18, sentAt: 19, notes: 20
};

const Q = {
  id: 'Student ID number',
  name: 'Your first and last name',
  cls: 'Your class',
  coachEmail: "Learning Coach's email address",
  title: 'Book title',
  author: 'Author',
  type: 'Was your book fiction or nonfiction?',
  q1: 'Question 1: Fiction: Who is the main character, and what are they like? Nonfiction: What is the main topic of the book?',
  q2: 'Question 2: Fiction: What problem happened in the story, and how was it solved? Nonfiction: What are two new facts you learned?',
  q3: 'Question 3: What was your favorite part of the book, and why?',
  rating: 'How many stars would you give this book?',
  recommend: 'Would you recommend this book to a friend?',
  science: 'Bonus: Did this book connect to anything we learned in science? Tell us how. (optional)'
};

/* ---------------- Menu ---------------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Book Club')
    .addItem('1. Set up (run once)', 'setup')
    .addItem('Send approved & queued now', 'processQueue')
    .addItem('Send a test certificate to me', 'sendTest')
    .addItem('Show form link', 'showFormLink')
    .addToUi();
}

/* ---------------- Setup ---------------- */

function setup() {
  const ui = SpreadsheetApp.getUi();
  const s = getSettings_();
  if (!s.TEACHER_EMAIL) {
    ui.alert('Fill in TEACHER_EMAIL on the Settings tab first, then run setup again.');
    return;
  }
  if (!s.CERT_TEMPLATE_ID || !s.COUPON_TEMPLATE_ID) {
    ui.alert('Paste the Google Slides IDs for the certificate and coupon templates on the Settings tab first (see the Start Here tab, step 2).');
    return;
  }

  // Drive folder for saved PDFs
  if (!s.PDF_FOLDER_ID) {
    const folder = DriveApp.createFolder((s.CLUB_NAME || 'Book Club') + ' - Certificates and Coupons');
    setSetting_('PDF_FOLDER_ID', folder.getId());
  }

  // Form
  let form;
  if (s.FORM_ID) {
    form = FormApp.openById(s.FORM_ID);
  } else {
    form = buildForm_(s);
    setSetting_('FORM_ID', form.getId());
  }
  setSetting_('FORM_URL', form.getPublishedUrl());
  setSetting_('FORM_EDIT_URL', form.getEditUrl());

  // Triggers (remove old ones first so setup can be re-run safely)
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onBookLogged').forForm(form).onFormSubmit().create();
  ScriptApp.newTrigger('processQueue').timeBased().everyHours(1).create();

  ui.alert('Setup complete!\n\nStudent form link (share this with students):\n' + form.getPublishedUrl() +
    '\n\nThe link is also saved on the Settings tab. Use "Send a test certificate to me" to see a sample.');
}

function buildForm_(s) {
  const club = s.CLUB_NAME || 'Book Club';
  const form = FormApp.create(club + ' - Book Log');
  form.setTitle(club + ' - Log a Book')
    .setDescription('Finished a book? Log it here! Answer every question in complete sentences. ' +
      'Every 5 books earns a Skip-a-Quiz Pass. Every 20 books earns a Reading Champion certificate! ' +
      'Rewards are emailed to your Learning Coach.')
    .setCollectEmail(false)
    .setProgressBar(true)
    .setConfirmationMessage('Great job! Your book was logged. Keep reading to earn your next reward!');

  const numberOnly = FormApp.createTextValidation().requireNumber()
    .setHelpText('Type your student ID number using numbers only.').build();
  form.addTextItem().setTitle(Q.id).setRequired(true).setValidation(numberOnly);
  form.addTextItem().setTitle(Q.name).setRequired(true);
  form.addListItem().setTitle(Q.cls).setChoiceValues(['TIH', 'TIN', 'TIR', 'TIU']).setRequired(true);
  const emailOnly = FormApp.createTextValidation().requireTextIsEmail()
    .setHelpText('Please enter a real email address.').build();
  form.addTextItem().setTitle(Q.coachEmail).setRequired(true).setValidation(emailOnly);

  form.addPageBreakItem().setTitle('About your book');
  form.addTextItem().setTitle(Q.title).setRequired(true);
  form.addTextItem().setTitle(Q.author).setRequired(true);
  form.addMultipleChoiceItem().setTitle(Q.type).setChoiceValues(['Fiction (a made-up story)', 'Nonfiction (true facts)']).setRequired(true);

  form.addPageBreakItem().setTitle('Show what you read')
    .setHelpText('Write at least 2 complete sentences for each question.');
  const min = Number(s.MIN_CHARS) || 40;
  const longEnough = () => FormApp.createParagraphTextValidation().requireTextLengthGreaterThanOrEqualTo(min)
    .setHelpText('Please write at least ' + min + ' characters (about 2 sentences).').build();
  form.addParagraphTextItem().setTitle(Q.q1).setRequired(true).setValidation(longEnough());
  form.addParagraphTextItem().setTitle(Q.q2).setRequired(true).setValidation(longEnough());
  form.addParagraphTextItem().setTitle(Q.q3).setRequired(true).setValidation(longEnough());
  form.addScaleItem().setTitle(Q.rating).setBounds(1, 5).setLabels('Not for me', 'Loved it!').setRequired(true);
  form.addMultipleChoiceItem().setTitle(Q.recommend).setChoiceValues(['Yes', 'Maybe', 'No']).setRequired(true);
  form.addParagraphTextItem().setTitle(Q.science).setRequired(false);
  return form;
}

/* ---------------- When a book is logged ---------------- */

function onBookLogged(e) {
  const a = {};
  e.response.getItemResponses().forEach(ir => { a[ir.getItem().getTitle()] = String(ir.getResponse()); });
  recordLog_(a);
  processQueue();
}

/** Saves one book log (from the Google Form or the website). Returns { status, bookNum, name }. */
function recordLog_(a) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const s = getSettings_();
    const id = String(a[Q.id] || '').replace(/\D/g, '');
    const student = findStudent_(id);
    const title = (a[Q.title] || '').trim();
    const log = sheet_(SHEETS.log);

    let status, notes = [];
    let name = student ? student.first + ' ' + student.last : (a[Q.name] || '').trim();
    if (!student) { status = 'Check ID'; notes.push('ID ' + id + ' is not on the Roster.'); }

    // duplicate check: same student, same book title already accepted
    if (student && isDuplicate_(log, id, title)) { status = 'Duplicate'; notes.push('This book was already logged by this student.'); }

    // quality check
    const min = Number(s.MIN_CHARS) || 40;
    const short = [Q.q1, Q.q2, Q.q3].filter(q => !looksReal_(a[q], min));
    if (!status && short.length) { status = 'Needs approval'; notes.push('One or more answers look too short or not real. Review before sending.'); }

    if (!status) status = (String(s.REVIEW_MODE).toLowerCase() === 'review') ? 'Needs approval' : 'Queued';

    const accepted = ['Queued', 'Needs approval'].indexOf(status) > -1 && !!student;
    const bookNum = accepted ? countAccepted_(log, id) + 1 : '';

    log.appendRow([
      new Date(), id ? Number(id) : '', name, student ? student.cls : (a[Q.cls] || ''),
      (a[Q.coachEmail] || '').trim(), title, (a[Q.author] || '').trim(), a[Q.type] || '',
      a[Q.q1] || '', a[Q.q2] || '', a[Q.q3] || '', a[Q.rating] || '', a[Q.recommend] || '', a[Q.science] || '',
      bookNum, '', status, '', '', notes.join(' ')
    ]);

    if (status === 'Check ID' || status === 'Duplicate' || status === 'Needs approval') {
      notifyTeacher_(s, status, name, id, title, notes.join(' '));
    }
    return { status: status, bookNum: bookNum, name: student ? student.first : '' };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- Website (GitHub Pages) connection ---------------- */
// Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone.
// Paste the Web app URL into the website's index.html (WEB_APP_URL).

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    if (d.website) return json_({ ok: false, message: 'Sorry, something went wrong.' }); // spam trap
    const s = getSettings_();
    const min = Number(s.MIN_CHARS) || 40;
    const req = ['id', 'name', 'cls', 'coachEmail', 'title', 'author', 'type', 'q1', 'q2', 'q3', 'rating', 'recommend'];
    const missing = req.filter(k => !String(d[k] || '').trim());
    if (missing.length) return json_({ ok: false, message: 'Please fill in every required question.' });
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(d.coachEmail).trim())) return json_({ ok: false, message: "Please check the Learning Coach's email address." });
    if (['q1', 'q2', 'q3'].some(k => String(d[k]).trim().length < min)) return json_({ ok: false, message: 'Each book question needs at least ' + min + ' characters (about 2 sentences).' });

    const a = {};
    Object.keys(Q).forEach(k => { a[Q[k]] = String(d[k] || '').slice(0, 3000); });
    const r = recordLog_(a);
    processQueue();
    const msgs = {
      'Queued': 'Great job' + (r.name ? ', ' + r.name : '') + '! That is book #' + r.bookNum + '. ' + nextGoals_(r.bookNum, s, getRewards_()),
      'Needs approval': 'Great job' + (r.name ? ', ' + r.name : '') + '! Your teacher will read your answers, then your book will count.',
      'Check ID': 'We could not find that student ID. Your teacher will check it and get back to you.',
      'Duplicate': 'You already logged this book. Try a new one to get closer to your next reward!'
    };
    return json_({ ok: r.status !== 'Check ID' && r.status !== 'Duplicate', status: r.status, message: msgs[r.status] || 'Your book was logged.' });
  } catch (err) {
    return json_({ ok: false, message: 'Sorry, something went wrong. Please try again later.' });
  }
}

/** Public stats for the website: class totals and rewards only, never student names. */
function doGet(e) {
  const s = getSettings_();
  const ro = sheet_(SHEETS.roster);
  const rows = ro.getRange(2, 1, ro.getLastRow() - 1, 5).getValues();
  const classes = {};
  rows.forEach(r => {
    const c = String(r[3]); if (!c) return;
    classes[c] = classes[c] || { students: 0, readers: 0, books: 0 };
    classes[c].students++;
    const n = Number(r[4]) || 0;
    classes[c].books += n; if (n > 0) classes[c].readers++;
  });
  const totalBooks = Object.keys(classes).reduce((t, c) => t + classes[c].books, 0);
  return json_({
    club: s.CLUB_NAME || 'Book Club', minChars: Number(s.MIN_CHARS) || 40,
    totalBooks: totalBooks, classes: classes,
    rewards: getRewards_().map(r => ({ every: r.every, title: r.title, text: r.text })),
    certificateEvery: Number(s.CERTIFICATE_EVERY) || 20
  });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- Sending ---------------- */

function processQueue() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    const s = getSettings_();
    const log = sheet_(SHEETS.log);
    const last = log.getLastRow();
    if (last < 2) return;
    const data = log.getRange(2, 1, last - 1, COL.notes).getValues();
    const rewards = getRewards_();

    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      const status = String(r[COL.status - 1]);
      const approved = String(r[COL.approve - 1]).trim().toLowerCase() === 'yes';
      const ready = status === 'Queued' || (status === 'Needs approval' && approved);
      if (!ready) continue;
      if (MailApp.getRemainingDailyQuota() < 1) {
        log.getRange(i + 2, COL.notes).setValue('Daily email limit reached; will send automatically later.');
        break;
      }
      const row = i + 2;
      try {
        const id = String(r[COL.id - 1]);
        const bookNum = Number(r[COL.bookNum - 1]) || countAccepted_(log, id);
        const reward = rewardFor_(rewards, bookNum, s);
        const certEvery = Number(s.CERTIFICATE_EVERY) || 20;
        const earnsCert = bookNum > 0 && bookNum % certEvery === 0;
        const info = {
          CLUB_NAME: s.CLUB_NAME || 'Book Club',
          STUDENT_NAME: String(r[COL.name - 1]),
          BOOK_TITLE: String(r[COL.title - 1]),
          AUTHOR: String(r[COL.author - 1]),
          BOOK_COUNT: String(bookNum),
          DATE: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM d, yyyy'),
          TEACHER_NAME: s.TEACHER_NAME || ''
        };
        const attachments = [];
        if (earnsCert) attachments.push(makePdf_(s.CERT_TEMPLATE_ID, info, 'Certificate - ' + info.STUDENT_NAME + ' - ' + bookNum + ' books', s));
        let couponTitle = '';
        if (reward) {
          const expires = new Date(); expires.setDate(expires.getDate() + (Number(reward.days) || 30));
          const cInfo = Object.assign({}, info, {
            COUPON_TITLE: reward.title, COUPON_TEXT: reward.text,
            CODE: 'CMS-' + id.slice(-4) + '-' + bookNum,
            EXPIRES: Utilities.formatDate(expires, Session.getScriptTimeZone(), 'MMMM d, yyyy')
          });
          attachments.push(makePdf_(s.COUPON_TEMPLATE_ID, cInfo, 'Coupon - ' + info.STUDENT_NAME + ' - ' + reward.title + ' - ' + bookNum + ' books', s));
          couponTitle = reward.title;
        }
        const progressEmails = String(s.PROGRESS_EMAILS || 'Yes').toLowerCase() !== 'no';
        if (attachments.length || progressEmails) {
          sendEmail_(s, String(r[COL.coachEmail - 1]), info, couponTitle, earnsCert, attachments, nextGoals_(bookNum, s, rewards));
        }
        log.getRange(row, COL.coupon).setValue([couponTitle, earnsCert ? 'Certificate' : ''].filter(Boolean).join(' + '));
        log.getRange(row, COL.status).setValue('Sent');
        log.getRange(row, COL.sentAt).setValue(new Date());
        log.getRange(row, COL.notes).setValue('');
      } catch (err) {
        log.getRange(row, COL.status).setValue('Error');
        log.getRange(row, COL.notes).setValue(String(err));
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function sendEmail_(s, to, info, couponTitle, earnsCert, attachments, nextGoal) {
  const club = info.CLUB_NAME;
  const n = info.BOOK_COUNT;
  let subject, prize = '';
  if (earnsCert && couponTitle) subject = info.STUDENT_NAME + ' read ' + n + ' books! Certificate and ' + couponTitle;
  else if (earnsCert) subject = info.STUDENT_NAME + ' read ' + n + ' books! Reading Champion certificate';
  else if (couponTitle) subject = info.STUDENT_NAME + ' earned a ' + couponTitle + '!';
  else subject = info.STUDENT_NAME + ' finished book #' + n + ' in the ' + club;
  if (earnsCert) prize += '<p>🏆 <b>' + esc_(info.STUDENT_NAME) + ' has read ' + n + ' books</b> and earned a <b>Reading Champion certificate</b>. It is attached. Please print it or share it with them!</p>';
  if (couponTitle) prize += '<p>🎉 They also earned a <b>' + esc_(couponTitle) + '</b>. The coupon is attached. They can show it to ' + esc_(s.TEACHER_NAME || 'their teacher') + ' to use it.</p>';
  const html =
    '<p>Hello,</p>' +
    '<p>Great news! <b>' + esc_(info.STUDENT_NAME) + '</b> just finished reading <b><i>' + esc_(info.BOOK_TITLE) +
    '</i></b> by ' + esc_(info.AUTHOR) + ' and answered the book questions. That makes <b>' + n +
    (n === '1' ? ' book' : ' books') + '</b> in the ' + esc_(club) + '.</p>' + prize +
    (nextGoal ? '<p>' + esc_(nextGoal) + '</p>' : '') +
    '<p>Keep reading!<br>' + esc_(s.TEACHER_NAME || '') + '<br>' + esc_(club) + '</p>';
  const opts = { htmlBody: html, name: club };
  if (attachments && attachments.length) opts.attachments = attachments;
  if (s.TEACHER_EMAIL && String(s.CC_TEACHER).toLowerCase() === 'yes') opts.cc = s.TEACHER_EMAIL;
  MailApp.sendEmail(to, subject, 'Great job reading! See the details in this email.', opts);
}

/** "Only 2 more books until the next Skip-a-Quiz Pass!" */
function nextGoals_(bookNum, s, rewards) {
  const parts = [];
  const couponEvery = rewards.length ? rewards[0].every : (Number(s.COUPON_EVERY) || 5);
  const certEvery = Number(s.CERTIFICATE_EVERY) || 20;
  if (couponEvery) {
    const left = couponEvery - (bookNum % couponEvery);
    parts.push(left + (left === 1 ? ' more book' : ' more books') + ' until the next ' + (rewards[0] ? rewards[0].title : 'reward coupon'));
  }
  if (certEvery) {
    const left = certEvery - (bookNum % certEvery);
    parts.push(left + (left === 1 ? ' more book' : ' more books') + ' until the next Reading Champion certificate');
  }
  return parts.length ? 'Next goals: ' + parts.join('; ') + '.' : '';
}

function notifyTeacher_(s, status, name, id, title, notes) {
  if (!s.TEACHER_EMAIL || MailApp.getRemainingDailyQuota() < 1) return;
  MailApp.sendEmail(s.TEACHER_EMAIL, 'Book Club: ' + status + ' - ' + name,
    'A book log needs your attention.\n\nStudent: ' + name + ' (ID ' + id + ')\nBook: ' + title +
    '\nStatus: ' + status + '\nNotes: ' + notes +
    '\n\nOpen the Book Log tab. To send it, type Yes in the Approve column, then choose Book Club > Send approved & queued now.');
}

function sendTest() {
  const s = getSettings_();
  const n = Number(s.CERTIFICATE_EVERY) || 20;
  const info = {
    CLUB_NAME: s.CLUB_NAME || 'Book Club', STUDENT_NAME: 'Sample Student', BOOK_TITLE: 'The Magic School Bus: Inside the Human Body',
    AUTHOR: 'Joanna Cole', BOOK_COUNT: String(n), DATE: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM d, yyyy'),
    TEACHER_NAME: s.TEACHER_NAME || ''
  };
  const rewards = getRewards_();
  const reward = rewards[0] || { title: 'Skip-a-Quiz Pass', text: 'Skip one quiz of your choice.', days: 30 };
  const cInfo = Object.assign({}, info, { COUPON_TITLE: reward.title, COUPON_TEXT: reward.text, CODE: 'CMS-0000-' + n, EXPIRES: 'next month' });
  const att = [makePdf_(s.CERT_TEMPLATE_ID, info, 'TEST Certificate', s), makePdf_(s.COUPON_TEMPLATE_ID, cInfo, 'TEST Coupon', s)];
  sendEmail_(s, s.TEACHER_EMAIL, info, reward.title, true, att, nextGoals_(n, s, rewards));
  SpreadsheetApp.getUi().alert('Test email sent to ' + s.TEACHER_EMAIL + '.');
}

function showFormLink() {
  const s = getSettings_();
  SpreadsheetApp.getUi().alert(s.FORM_URL ? 'Student form link:\n' + s.FORM_URL : 'Run setup first.');
}

/* ---------------- PDFs ---------------- */

function makePdf_(templateId, values, fileName, s) {
  const folder = s.PDF_FOLDER_ID ? DriveApp.getFolderById(s.PDF_FOLDER_ID) : DriveApp.getRootFolder();
  const copy = DriveApp.getFileById(templateId).makeCopy(fileName, folder);
  const pres = SlidesApp.openById(copy.getId());
  Object.keys(values).forEach(k => pres.replaceAllText('{{' + k + '}}', String(values[k] || '')));
  // Shrink long book titles so they stay on the certificate
  if (values.BOOK_TITLE && values.BOOK_TITLE.length > 34) shrinkText_(pres, values.BOOK_TITLE, values.BOOK_TITLE.length > 60 ? 12 : 15);
  if (values.STUDENT_NAME && values.STUDENT_NAME.length > 24) shrinkText_(pres, values.STUDENT_NAME, 24);
  pres.saveAndClose();
  const pdf = copy.getAs('application/pdf').setName(fileName + '.pdf');
  if (String(s.SAVE_PDFS).toLowerCase() === 'yes') folder.createFile(pdf);
  copy.setTrashed(true);
  return pdf;
}

function shrinkText_(pres, text, size) {
  pres.getSlides().forEach(slide => slide.getShapes().forEach(shape => {
    try {
      const tr = shape.getText();
      if (tr.asString().trim() === text.trim()) tr.getTextStyle().setFontSize(size);
    } catch (e) { /* shape has no text */ }
  }));
}

/* ---------------- Helpers ---------------- */

function sheet_(name) { return SpreadsheetApp.getActive().getSheetByName(name); }

function getSettings_() {
  const sh = sheet_(SHEETS.settings);
  const vals = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 2).getValues();
  const out = {};
  vals.forEach(([k, v]) => { if (k) out[String(k).trim()] = String(v).trim(); });
  return out;
}

function setSetting_(key, value) {
  const sh = sheet_(SHEETS.settings);
  const keys = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(r => String(r[0]).trim());
  const i = keys.indexOf(key);
  if (i > -1) sh.getRange(i + 2, 2).setValue(value);
  else sh.appendRow([key, value]);
}

function findStudent_(id) {
  if (!id) return null;
  const sh = sheet_(SHEETS.roster);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  for (const r of rows) if (String(r[0]).replace(/\D/g, '') === id) return { id: id, first: r[1], last: r[2], cls: r[3] };
  return null;
}

function countAccepted_(log, id) {
  const last = log.getLastRow();
  if (last < 2) return 0;
  return log.getRange(2, 1, last - 1, COL.status).getValues()
    .filter(r => String(r[COL.id - 1]) === String(id) && ['Queued', 'Needs approval', 'Sent'].indexOf(String(r[COL.status - 1])) > -1).length;
}

function isDuplicate_(log, id, title) {
  const last = log.getLastRow();
  if (last < 2) return false;
  const t = norm_(title);
  return log.getRange(2, 1, last - 1, COL.status).getValues()
    .some(r => String(r[COL.id - 1]) === String(id) && norm_(r[COL.title - 1]) === t &&
      ['Queued', 'Needs approval', 'Sent'].indexOf(String(r[COL.status - 1])) > -1);
}

function looksReal_(text, min) {
  const t = String(text || '').trim();
  if (t.length < min) return false;
  const words = t.split(/\s+/).filter(Boolean);
  if (words.length < 6) return false;
  if (/(.)\1{5,}/.test(t)) return false;                 // aaaaaaa
  if (new Set(words.map(w => w.toLowerCase())).size < 4) return false; // same word repeated
  return true;
}

/** Rewards tab: Every (books) | Coupon Title | Coupon Text | Good For (days). */
function getRewards_() {
  const sh = sheet_(SHEETS.rewards);
  const rows = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 4).getValues();
  return rows.filter(r => Number(r[0]) > 0 && r[1] !== '')
    .map(r => ({ every: Number(r[0]), title: String(r[1]), text: String(r[2]), days: Number(r[3]) || 30 }))
    .sort((a, b) => a.every - b.every);
}

/** The coupon earned at this book count, if any (largest matching "every" wins). */
function rewardFor_(rewards, bookNum, s) {
  if (!bookNum) return null;
  const hits = rewards.filter(r => bookNum % r.every === 0);
  return hits.length ? hits[hits.length - 1] : null;
}

function norm_(t) { return String(t || '').toLowerCase().replace(/[^a-z0-9]/g, ''); }
function esc_(t) { return String(t || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
