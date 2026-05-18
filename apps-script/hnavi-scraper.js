// Flow:
// login -> get jobs/projects from two URLs -> loop through all job detail pages found on each listing page
// -> extract fields -> compare against existing sheet data -> append only new rows
// -> send one webhook notification per new record.

// 処理の流れ:
// ログイン -> 案件一覧を取得 -> 各案件の詳細ページを処理
// -> 必要項目を抽出 -> 既存シートの job_url と比較
// -> 新規案件のみ追記 -> 新規案件ごとに Cloudflare Worker へ通知

// MAIN FUNCTIONS
function scrapeHnaviJobs() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error('Another scrape is already running. Skipping this run.');
  }

  try {
    scrapeHnaviJobs_();
  } finally {
    lock.releaseLock();
  }
}

function scrapeHnaviJobs_() {
  const LOGIN_URL = 'https://developer.hnavi.co.jp/developer_users/sign_in';
  const JOBS_URL = 'https://developer.hnavi.co.jp/jobs';
  const SAAS_JOBS_URL = 'https://developer.hnavi.co.jp/jobs?saas=saas';

  const scriptProperties = PropertiesService.getScriptProperties();

  const PASSWORD = scriptProperties.getProperty('PASSWORD');
  const EMAIL = scriptProperties.getProperty('EMAIL');
  const SHEET = scriptProperties.getProperty('SHEET');

  if (!EMAIL || !PASSWORD) {
    throw new Error('EMAIL or PASSWORD is not set in Script Properties.');
  }

  if (!SHEET) {
    throw new Error('SHEET is not set in Script Properties.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Open this script from a Google Sheet: Extensions → Apps Script.');
  }

  const session = createSession();

  const loginPage = session.fetch(LOGIN_URL, {
    method: 'get',
    followRedirects: true,
    muteHttpExceptions: true
  });

  const loginHtml = loginPage.getContentText('UTF-8');

  const authenticityToken =
    extractInputValue(loginHtml, 'authenticity_token') ||
    extractMetaCsrfToken(loginHtml);

  if (!authenticityToken) {
    writeDebugToSheet_({
      step: 'open login page',
      loginPageStatus: loginPage.getResponseCode(),
      loginPageTitle: extractTitle(loginHtml),
      loginHtmlSnippet: loginHtml.substring(0, 3000)
    });

    throw new Error('Could not find authenticity_token. Debug details written to "debug" sheet.');
  }

  const formAction = extractFormAction(loginHtml) || '/developer_users/sign_in';
  const loginPostUrl = absolutizeUrl(formAction, LOGIN_URL);

  const emailField = findInputName(loginHtml, 'email') || 'developer_user[email]';
  const passwordField = findInputName(loginHtml, 'password') || 'developer_user[password]';

  const payload = {};
  payload['authenticity_token'] = authenticityToken;
  payload[emailField] = EMAIL;
  payload[passwordField] = PASSWORD;
  payload['developer_user[remember_me]'] = '0';
  payload['commit'] = 'ログイン';

  const loginResponse = session.fetch(loginPostUrl, {
    method: 'post',
    followRedirects: false,
    muteHttpExceptions: true,
    payload: payload
  });

  const redirectLocation = getHeader(loginResponse, 'Location');
  if (redirectLocation) {
    session.fetch(absolutizeUrl(redirectLocation, loginPostUrl), {
      method: 'get',
      followRedirects: true,
      muteHttpExceptions: true
    });
  }

  const devNewRows = scrapeJobsPageToSheet_({
    session: session,
    ss: ss,
    jobsUrl: JOBS_URL,
    sheetName: 'dev-projects',
    loginStatus: loginResponse.getResponseCode(),
    loginRedirect: redirectLocation
  });

  const saasNewRows = scrapeJobsPageToSheet_({
    session: session,
    ss: ss,
    jobsUrl: SAAS_JOBS_URL,
    sheetName: 'saas-projects',
    loginStatus: loginResponse.getResponseCode(),
    loginRedirect: redirectLocation
  });

  notifyNewJobs_(SHEET, 'dev-projects', devNewRows);
  notifyNewJobs_(SHEET, 'saas-projects', saasNewRows);

  Logger.log('New normal jobs: ' + devNewRows.length);
  Logger.log('New SaaS jobs: ' + saasNewRows.length);
}


// HELPERS
/**
 * Fetches one jobs listing page, extracts job details, compares with existing sheet rows,
 * appends only new rows, and returns the newly appended rows.
 */
function scrapeJobsPageToSheet_(config) {
  const session = config.session;
  const ss = config.ss;
  const jobsUrl = config.jobsUrl;
  const sheetName = config.sheetName;
  const loginStatus = config.loginStatus;
  const loginRedirect = config.loginRedirect;

  const jobsResponse = session.fetch(jobsUrl, {
    method: 'get',
    followRedirects: true,
    muteHttpExceptions: true
  });

  const jobsHtml = jobsResponse.getContentText('UTF-8');

  if (looksLikeLoginPage(jobsHtml)) {
    writeDebugToSheet_({
      step: 'fetch jobs page after login',
      loginStatus: loginStatus,
      loginRedirect: loginRedirect,
      jobsStatus: jobsResponse.getResponseCode(),
      jobsFinalTitle: extractTitle(jobsHtml),
      jobsUrl: jobsUrl,
      jobsHtmlSnippet: jobsHtml.substring(0, 3000)
    });

    throw new Error('Login appears to have failed while fetching ' + jobsUrl + '. Debug details were written to the "debug" sheet.');
  }

  const jobs = extractJobsFromJobsPage(jobsHtml, jobsUrl);

  if (jobs.length === 0) {
    writeDebugToSheet_({
      step: 'extract jobs from jobs page',
      jobsStatus: jobsResponse.getResponseCode(),
      jobsFinalTitle: extractTitle(jobsHtml),
      jobsUrl: jobsUrl,
      jobsHtmlSnippet: jobsHtml.substring(0, 3000)
    });

    throw new Error('Could not find job links on ' + jobsUrl + '. Debug details written to "debug" sheet.');
  }

  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  ensureJobsSheetHeader_(sheet);

  const existingUrls = getExistingJobUrls_(sheet);
  const rows = [];

  for (const job of jobs) {
    if (!job.url) continue;

    // Already in sheet, so skip detail fetch and notification.
    if (existingUrls.has(job.url)) {
      continue;
    }

    const detailResponse = session.fetch(job.url, {
      method: 'get',
      followRedirects: true,
      muteHttpExceptions: true
    });

    const detailHtml = detailResponse.getContentText('UTF-8');

    if (looksLikeLoginPage(detailHtml)) {
      writeDebugToSheet_({
        step: 'fetch job detail page',
        detailStatus: detailResponse.getResponseCode(),
        detailUrl: job.url,
        detailFinalTitle: extractTitle(detailHtml),
        detailHtmlSnippet: detailHtml.substring(0, 3000)
      });

      throw new Error('A job detail page showed login page. Debug details written to "debug" sheet.');
    }

    const projectTitle =
      extractJobTitleFromDetail(detailHtml) ||
      job.title ||
      '';

    const entryConditions = extractEntryConditions(detailHtml);
    const inquiryContent = extractPreWrapAfterHeading(detailHtml, 'お問い合わせ時の内容');
    const hearingContent = extractPreWrapAfterHeading(detailHtml, '発注ナビ担当者のヒアリング内容');

    const missing = [];
    if (!entryConditions) missing.push('エントリー条件 not found');
    if (!inquiryContent) missing.push('お問い合わせ時の内容 not found');
    if (!hearingContent) missing.push('発注ナビ担当者のヒアリング内容 not found');

    const row = [
      job.url,
      projectTitle,
      entryConditions,
      inquiryContent,
      hearingContent,
      new Date(),
      missing.length ? 'PARTIAL' : 'OK',
      missing.join(' / ')
    ];

    rows.push(row);

    // Prevent duplicate insertion within the same run if the listing contains duplicates.
    existingUrls.add(job.url);

    Utilities.sleep(1000);
  }

 if (rows.length > 0) {
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, 8).setValues(rows);
  sheet.autoResizeColumns(1, 8);

  Logger.log('Appended ' + rows.length + ' new rows to ' + sheetName);

  rows.forEach(row => {
    Logger.log('Appended row: ' + row[1] + ' | ' + row[0]);
  });
} else {
  Logger.log('No rows appended to ' + sheetName);
}

return rows;
}

/**
 * Ensures the sheet has the expected header row.
 */
function ensureJobsSheetHeader_(sheet) {
  const headers = [
    'job_url',
    'project_title',
    'エントリー条件',
    'お問い合わせ時の内容',
    '発注ナビ担当者のヒアリング内容',
    'scraped_at',
    'status',
    'note'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  const currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const needsHeader = currentHeaders[0] !== 'job_url';

  if (needsHeader) {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

/**
 * Reads existing job_url values from column A.
 */
function getExistingJobUrls_(sheet) {
  const urls = new Set();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return urls;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  values.forEach(row => {
    const url = String(row[0] || '').trim();
    if (url) urls.add(url);
  });

  return urls;
}

/**
 * Sends one notification per new job.
 */
function notifyNewJobs_(sheetUrl, sheetName, newRows) {
  if (!newRows || newRows.length === 0) {
    Logger.log('No new jobs for ' + sheetName);
    return;
  }

  Logger.log('New jobs for ' + sheetName + ': ' + newRows.length);

  newRows.forEach(row => {
    Logger.log('New job detected: ' + row[1] + ' | ' + row[0]);

    notifyCloudflareNewJob_({
      sheetUrl: sheetUrl,
      sheetName: sheetName,
      jobUrl: row[0],
      projectTitle: row[1],
      entryConditions: row[2],
      inquiryContent: row[3],
      hearingContent: row[4],
      scrapedAt: row[5],
      status: row[6],
      note: row[7]
    });

    Utilities.sleep(300);
  });
}

/**
 * Sends a single new-job event to Cloudflare Worker.
 */
function notifyCloudflareNewJob_(job) {
  const cloudflareUrl = PropertiesService
    .getScriptProperties()
    .getProperty('CLOUDFLARE_WORKER_URL');

  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty('CLOUDFLARE_API_KEY');

  if (!cloudflareUrl) {
    throw new Error('CLOUDFLARE_WORKER_URL is not set in Script Properties.');
  }

  if (!apiKey) {
    throw new Error('CLOUDFLARE_API_KEY is not set in Script Properties.');
  }

  const payload = {
    eventType: 'newJob',
    sheetUrl: job.sheetUrl,
    sheetName: job.sheetName,
    jobUrl: job.jobUrl,
    projectTitle: job.projectTitle,
    entryConditions: job.entryConditions,
    inquiryContent: job.inquiryContent,
    hearingContent: job.hearingContent,
    scrapedAt: job.scrapedAt,
    status: job.status,
    note: job.note
  };

  const response = UrlFetchApp.fetch(cloudflareUrl, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  const body = response.getContentText();

  Logger.log('Cloudflare notification status: ' + status);
  Logger.log(body);

  if (status < 200 || status >= 300) {
    throw new Error('Cloudflare new job notification failed: ' + status + ' ' + body);
  }
}

/**
 * Extracts the エントリー条件 block.
 */
function extractEntryConditions(html) {
  const source = String(html || '');

  const startIndex = source.indexOf('エントリー条件');
  if (startIndex < 0) return '';

  const stopHeadings = [
    'お問い合わせ時の内容',
    '発注ナビ担当者のヒアリング内容',
    '予算',
    '納期',
    'カテゴリ'
  ];

  let endIndex = source.length;

  stopHeadings.forEach(heading => {
    const idx = source.indexOf(heading, startIndex + 'エントリー条件'.length);
    if (idx >= 0 && idx < endIndex) {
      endIndex = idx;
    }
  });

  const sectionHtml = source.substring(startIndex, endIndex);

  let visible = htmlToVisibleTextForEntryConditions(sectionHtml);

  visible = visible
    .replace(/^エントリー条件\s*/g, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();

  const labels = [
    '拠点指定',
    'スキル・その他',
    '開発言語',
    'フレームワーク',
    'データベース',
    'インフラ',
    '対応工程',
    '業種',
    'エリア'
  ];

  const foundLabels = [];

  labels.forEach(label => {
    const idx = visible.indexOf(label);
    if (idx >= 0) {
      foundLabels.push({
        label: label,
        index: idx
      });
    }
  });

  foundLabels.sort((a, b) => a.index - b.index);

  if (foundLabels.length === 0) {
    return visible;
  }

  const parts = [];

  for (let i = 0; i < foundLabels.length; i++) {
    const current = foundLabels[i];
    const next = foundLabels[i + 1];

    const valueStart = current.index + current.label.length;
    const valueEnd = next ? next.index : visible.length;

    const value = visible
      .substring(valueStart, valueEnd)
      .replace(/\n{3,}/g, '\n\n')
      .trim();

    if (value) {
      parts.push(current.label + ':\n' + value);
    }
  }

  return parts.join('\n\n').trim();
}

/**
 * Converts the エントリー条件 HTML block into readable text.
 */
function htmlToVisibleTextForEntryConditions(html) {
  return decodeHtml(String(html || ''))
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(
      /<div[^>]+class=['"][^'"]*circle-char[^'"]*['"][^>]*>([\s\S]*?)<\/div>/gi,
      '\n$1. '
    )
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(div|p|li|tr|section|article|h1|h2|h3|td|th|dd|dt)>/gi, '\n')
    .replace(/<svg[\s\S]*?<\/svg>/gi, '')
    .replace(/<img[^>]*>/gi, '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\r/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .replace(/[ \t]*\n[ \t]*/g, '\n')
    .replace(/\n\s*(\d+)\s*\.\s*/g, '\n$1. ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/**
 * Extracts the next .pre-wrap block after a visible heading.
 */
function extractPreWrapAfterHeading(html, heading) {
  const source = String(html || '');
  const headingIndex = findHeadingIndex(source, heading);
  if (headingIndex < 0) return '';

  const afterHeading = source.substring(headingIndex);

  const preWrapMatch = afterHeading.match(
    /<div[^>]+class=['"][^'"]*\bpre-wrap\b[^'"]*['"][^>]*>([\s\S]*?)<\/div>/i
  );

  return preWrapMatch ? cleanTextPreserveLines(preWrapMatch[1]) : '';
}

/**
 * Finds a div heading whose visible text contains the target heading.
 */
function findHeadingIndex(html, heading) {
  const source = String(html || '');
  const divRegex = /<div[^>]*>[\s\S]*?<\/div>/gi;
  let match;

  while ((match = divRegex.exec(source)) !== null) {
    const text = cleanText(match[0]);
    if (text === heading || text.indexOf(heading) >= 0) {
      return match.index;
    }
  }

  return -1;
}

/**
 * Simple cookie/session wrapper for UrlFetchApp.
 */
function createSession() {
  const jar = {};

  function updateCookies(response) {
    const headers = response.getAllHeaders();
    const setCookie = headers['Set-Cookie'] || headers['set-cookie'];

    if (!setCookie) return;

    const cookies = Array.isArray(setCookie) ? setCookie : [setCookie];

    cookies.forEach(cookieText => {
      const firstPart = String(cookieText).split(';')[0];
      const eqIndex = firstPart.indexOf('=');

      if (eqIndex > 0) {
        const name = firstPart.substring(0, eqIndex).trim();
        const value = firstPart.substring(eqIndex + 1).trim();
        jar[name] = value;
      }
    });
  }

  function cookieHeader() {
    return Object.keys(jar)
      .map(name => name + '=' + jar[name])
      .join('; ');
  }

  return {
    fetch: function(url, options) {
      options = options || {};
      options.headers = options.headers || {};

      if (cookieHeader()) {
        options.headers.Cookie = cookieHeader();
      }

      options.headers['User-Agent'] =
        options.headers['User-Agent'] ||
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36';

      options.headers['Accept'] =
        options.headers['Accept'] ||
        'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8';

      const response = UrlFetchApp.fetch(url, options);
      updateCookies(response);
      return response;
    }
  };
}

/**
 * Extracts all job cards from the listing page.
 */
function extractJobsFromJobsPage(html, baseUrl) {
  const jobs = [];
  const seenUrls = new Set();

  const cardRegex = /<a[^>]+class=['"][^'"]*card-body[^'"]*['"][^>]+href=['"]([^'"]+)['"][^>]*>([\s\S]*?)<\/a>/gi;

  let match;

  while ((match = cardRegex.exec(String(html || ''))) !== null) {
    const href = match[1];
    const cardHtml = match[2];

    const url = absolutizeUrl(href, baseUrl || 'https://developer.hnavi.co.jp/jobs');

    if (!url || seenUrls.has(url)) {
      continue;
    }

    seenUrls.add(url);

    const titleMatch = cardHtml.match(
      /<div[^>]+class=['"][^'"]*\btitle\b[^'"]*['"][^>]*>([\s\S]*?)<\/div>/i
    );

    const title = titleMatch ? cleanText(titleMatch[1]) : '';

    jobs.push({
      url: url,
      title: title
    });
  }

  return jobs;
}

function extractJobTitleFromDetail(html) {
  const candidates = [
    /<div[^>]+class=['"][^'"]*\btitle\b[^'"]*['"][^>]*>([\s\S]*?)<\/div>/i,
    /<h1[^>]*>([\s\S]*?)<\/h1>/i,
    /<h2[^>]*>([\s\S]*?)<\/h2>/i,
    /<div[^>]+class=['"][^'"]*job-title[^'"]*['"][^>]*>([\s\S]*?)<\/div>/i,
    /<title[^>]*>([\s\S]*?)<\/title>/i
  ];

  for (const regex of candidates) {
    const match = String(html || '').match(regex);
    if (match) {
      const text = cleanText(match[1]);
      if (text && !text.includes('発注ナビ-開発会社様用画面')) {
        return text;
      }
    }
  }

  return '';
}

function extractMetaCsrfToken(html) {
  const match = String(html || '').match(
    /<meta[^>]+name=['"]csrf-token['"][^>]+content=['"]([^'"]+)['"]/i
  );

  return match ? decodeHtml(match[1]) : '';
}

function extractInputValue(html, inputName) {
  const inputRegex = /<input[^>]*>/gi;
  let match;

  while ((match = inputRegex.exec(String(html || ''))) !== null) {
    const tag = match[0];
    const nameMatch = tag.match(/name=['"]([^'"]+)['"]/i);

    if (!nameMatch) continue;

    if (nameMatch[1] === inputName) {
      const valueMatch = tag.match(/value=['"]([^'"]*)['"]/i);
      return valueMatch ? decodeHtml(valueMatch[1]) : '';
    }
  }

  return '';
}

function extractFormAction(html) {
  const formRegex = /<form[^>]*method=['"]post['"][^>]*>|<form[^>]*>/gi;
  let match;

  while ((match = formRegex.exec(String(html || ''))) !== null) {
    const tag = match[0];

    if (!/method=['"]post['"]/i.test(tag)) continue;

    const actionMatch = tag.match(/action=['"]([^'"]+)['"]/i);
    if (actionMatch) return decodeHtml(actionMatch[1]);
  }

  const fallback = String(html || '').match(/<form[^>]+action=['"]([^'"]+)['"][^>]*>/i);
  return fallback ? decodeHtml(fallback[1]) : '';
}

function findInputName(html, typeOrNameHint) {
  const inputRegex = /<input[^>]*>/gi;
  let match;

  while ((match = inputRegex.exec(String(html || ''))) !== null) {
    const tag = match[0];
    const typeMatch = tag.match(/type=['"]([^'"]+)['"]/i);
    const nameMatch = tag.match(/name=['"]([^'"]+)['"]/i);

    if (!nameMatch) continue;

    const type = typeMatch ? typeMatch[1].toLowerCase() : '';
    const name = nameMatch[1];

    if (typeOrNameHint === 'email') {
      if (type === 'email' || /email/i.test(name)) return name;
    }

    if (typeOrNameHint === 'password') {
      if (type === 'password' || /password/i.test(name)) return name;
    }
  }

  return '';
}

function looksLikeLoginPage(html) {
  const text = String(html || '');

  return /ログインしてください/.test(text) ||
    (
      /password/i.test(text) &&
      /authenticity_token/i.test(text) &&
      /ログイン|sign_in|Sign in|開発会社用ログイン/i.test(text)
    );
}

function writeDebugToSheet_(debug) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('debug') || ss.insertSheet('debug');

  sheet.clearContents();

  const rows = [
    ['field', 'value'],
    ['step', debug.step || ''],
    ['login_page_status', debug.loginPageStatus || ''],
    ['login_page_title', debug.loginPageTitle || ''],
    ['login_status', debug.loginStatus || ''],
    ['login_redirect', debug.loginRedirect || ''],
    ['jobs_status', debug.jobsStatus || ''],
    ['jobs_final_title', debug.jobsFinalTitle || ''],
    ['jobs_url', debug.jobsUrl || ''],
    ['detail_status', debug.detailStatus || ''],
    ['detail_url', debug.detailUrl || ''],
    ['detail_final_title', debug.detailFinalTitle || ''],
    ['login_html_snippet', debug.loginHtmlSnippet || ''],
    ['jobs_html_snippet', debug.jobsHtmlSnippet || ''],
    ['detail_html_snippet', debug.detailHtmlSnippet || '']
  ];

  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
}

function extractTitle(html) {
  const match = String(html || '').match(/<title[^>]*>([\s\S]*?)<\/title>/i);
  return match ? cleanText(match[1]) : '';
}

function cleanText(value) {
  return decodeHtml(String(value || ''))
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(div|p|li|tr|section|article|h1|h2|h3|td|th|dd|dt)>/gi, '\n')
    .replace(/<[^>]*>/g, ' ')
    .replace(/\r/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n\s+/g, '\n')
    .replace(/\n{2,}/g, '\n')
    .trim();
}

function cleanTextPreserveLines(value) {
  return decodeHtml(String(value || ''))
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(div|p|li|tr|section|article|h1|h2|h3|td|th|dd|dt)>/gi, '\n')
    .replace(/<[^>]*>/g, ' ')
    .replace(/\r/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .replace(/[ \t]*\n[ \t]*/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function decodeHtml(value) {
  return String(value || '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&#x2F;/g, '/')
    .replace(/&#x27;/g, "'")
    .replace(/&#x60;/g, '`');
}

function absolutizeUrl(url, baseUrl) {
  if (!url) return '';
  if (/^https?:\/\//i.test(url)) return url;

  const base = String(baseUrl || '').match(/^(https?:\/\/[^\/]+)/i);
  const origin = base ? base[1] : '';

  if (url.startsWith('/')) return origin + url;

  return origin + '/' + url;
}

function getHeader(response, name) {
  const headers = response.getAllHeaders();
  return headers[name] || headers[name.toLowerCase()] || '';
}

// POLLING
/**
 * Optional helper: run once manually to create polling trigger.
 * Adjust everyMinutes(15), everyHours(1) etc. as needed.
 */
function createPollingTrigger() {
  deletePollingTriggers();

  ScriptApp.newTrigger('scrapeHnaviJobs')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('Hourly polling trigger created.');
}

/**
 * Optional helper: remove existing triggers for this scraper.
 */
function deletePollingTriggers() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'scrapeHnaviJobs') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  Logger.log('Existing scraper triggers deleted.');
}

function listPollingTriggers() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    Logger.log(
      'Function: ' +
      trigger.getHandlerFunction() +
      ', Event type: ' +
      trigger.getEventType()
    );
  });
}