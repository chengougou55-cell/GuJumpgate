(function outlookEmailUtilsModule(root, factory) {
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = factory();
    return;
  }

  root.OutlookEmailUtils = factory();
})(typeof self !== 'undefined' ? self : globalThis, function createOutlookEmailUtils() {
  const DEFAULT_BASE_URL = 'http://156.239.40.207:15000';
  const DEFAULT_API_KEY = 'ad9f6283fdee7ad92e6b2adef5a45050';
  const DEFAULT_MAIL_PAGE_SIZE = 20;

  function firstNonEmptyString(values) {
    for (const value of values) {
      if (value === undefined || value === null) continue;
      const normalized = String(value).trim();
      if (normalized) return normalized;
    }
    return '';
  }

  function normalizeOutlookEmailBaseUrl(rawValue = '') {
    const value = String(rawValue || '').trim();
    if (!value) return '';

    const candidate = /^[a-zA-Z][a-zA-Z\d+\-.]*:\/\//.test(value) ? value : `https://${value}`;
    try {
      const parsed = new URL(candidate);
      parsed.hash = '';
      parsed.search = '';
      const pathname = parsed.pathname === '/' ? '' : parsed.pathname.replace(/\/+$/, '');
      return `${parsed.origin}${pathname}`;
    } catch {
      return '';
    }
  }

  function joinOutlookEmailUrl(baseUrl, path) {
    const normalizedBase = normalizeOutlookEmailBaseUrl(baseUrl);
    const normalizedPath = String(path || '').trim();
    if (!normalizedBase || !normalizedPath) return normalizedBase || '';
    return `${normalizedBase}${normalizedPath.startsWith('/') ? '' : '/'}${normalizedPath}`;
  }

  function buildOutlookEmailHeaders(config = {}, options = {}) {
    const headers = {};
    const apiKey = firstNonEmptyString([
      config.apiKey,
      config.outlookEmailApiKey,
      options.apiKey,
    ]);
    if (apiKey) {
      headers['X-API-Key'] = apiKey;
    }
    if (options.json) {
      headers['Content-Type'] = 'application/json';
    }
    if (options.acceptJson !== false) {
      headers.Accept = 'application/json';
    }
    return headers;
  }

  function normalizeOutlookEmailAddress(value = '') {
    return String(value || '').trim().toLowerCase();
  }

  function getOutlookEmailAccountRows(payload) {
    if (Array.isArray(payload)) return payload;
    if (!payload || typeof payload !== 'object') return [];

    const candidates = [
      payload.accounts,
      payload.items,
      payload.list,
      payload.rows,
      payload.records,
      payload.data,
      payload?.data?.accounts,
      payload?.data?.items,
      payload?.data?.list,
      payload?.data?.rows,
    ];

    for (const candidate of candidates) {
      if (Array.isArray(candidate)) return candidate;
    }
    return [];
  }

  function normalizeOutlookEmailAccount(row = {}) {
    if (!row || typeof row !== 'object') return null;
    const email = normalizeOutlookEmailAddress(firstNonEmptyString([
      row.email,
      row.address,
      row.mail,
      row.account,
    ]));
    if (!email) return null;

    return {
      id: firstNonEmptyString([row.id, row.account_id, row.accountId, email]),
      email,
      status: firstNonEmptyString([row.status, row.last_refresh_status]),
      provider: firstNonEmptyString([row.provider, row.account_type]),
      groupId: firstNonEmptyString([row.group_id, row.groupId]),
      groupName: firstNonEmptyString([row.group_name, row.groupName]),
      remark: firstNonEmptyString([row.remark, row.note]),
      createdAt: firstNonEmptyString([row.created_at, row.createdAt]),
      updatedAt: firstNonEmptyString([row.updated_at, row.updatedAt]),
      raw: row,
    };
  }

  function normalizeOutlookEmailAccounts(payload) {
    return getOutlookEmailAccountRows(payload)
      .map((row) => normalizeOutlookEmailAccount(row))
      .filter(Boolean);
  }

  function stripHtmlTags(value = '') {
    return String(value || '')
      .replace(/<style[\s\S]*?<\/style>/gi, ' ')
      .replace(/<script[\s\S]*?<\/script>/gi, ' ')
      .replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/gi, ' ')
      .replace(/&amp;/gi, '&')
      .replace(/&lt;/gi, '<')
      .replace(/&gt;/gi, '>')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function getOutlookEmailMailRows(payload) {
    if (Array.isArray(payload)) return payload;
    if (!payload || typeof payload !== 'object') return [];

    const candidates = [
      payload.emails,
      payload.messages,
      payload.items,
      payload.list,
      payload.rows,
      payload.records,
      payload.data,
      payload?.data?.emails,
      payload?.data?.messages,
      payload?.data?.items,
      payload?.data?.list,
      payload?.data?.rows,
    ];

    for (const candidate of candidates) {
      if (Array.isArray(candidate)) return candidate;
    }
    return [];
  }

  function normalizeOutlookEmailReceivedDateTime(value) {
    if (!value && value !== 0) return '';
    if (typeof value === 'number' && Number.isFinite(value)) {
      const timestamp = value > 0 && value < 100000000000 ? value * 1000 : value;
      return new Date(timestamp).toISOString();
    }
    const source = String(value || '').trim();
    if (!source) return '';
    if (/^\d+$/.test(source)) {
      const numeric = Number(source);
      if (Number.isFinite(numeric)) {
        const timestamp = numeric > 0 && numeric < 100000000000 ? numeric * 1000 : numeric;
        return new Date(timestamp).toISOString();
      }
    }
    const parsed = Date.parse(source);
    return Number.isFinite(parsed) ? new Date(parsed).toISOString() : source;
  }

  function normalizeOutlookEmailMessage(row = {}, fallbackFolder = '') {
    if (!row || typeof row !== 'object') return null;

    const fromAddress = firstNonEmptyString([
      row.from,
      row.sender,
      row.from_email,
      row.sender_email,
      row?.from?.emailAddress?.address,
      row?.from?.EmailAddress?.Address,
    ]);
    const htmlContent = firstNonEmptyString([row.html, row.body_html, row.bodyHtml]);
    const textContent = firstNonEmptyString([
      row.body_preview,
      row.bodyPreview,
      row.preview,
      row.text,
      row.plainText,
      row.content_text,
      row.body,
      row.content,
    ]);
    const bodyPreview = stripHtmlTags(textContent || htmlContent);
    const folder = firstNonEmptyString([row.folder, row.mailbox, fallbackFolder]);

    return {
      id: firstNonEmptyString([row.id, row.message_id, row.messageId, row.internetMessageId]),
      address: normalizeOutlookEmailAddress(firstNonEmptyString([row.to, row.email, row.recipient])),
      subject: firstNonEmptyString([row.subject, row.title]),
      from: {
        emailAddress: {
          address: fromAddress,
        },
      },
      bodyPreview,
      raw: htmlContent || textContent || '',
      folder,
      mailbox: folder,
      receivedDateTime: normalizeOutlookEmailReceivedDateTime(firstNonEmptyString([
        row.receivedDateTime,
        row.received_date_time,
        row.received_at,
        row.date,
        row.created_at,
        row.createdAt,
      ])),
    };
  }

  function normalizeOutlookEmailMailApiMessages(payload, fallbackFolder = '') {
    return getOutlookEmailMailRows(payload)
      .map((row) => normalizeOutlookEmailMessage(row, fallbackFolder))
      .filter(Boolean);
  }

  return {
    DEFAULT_API_KEY,
    DEFAULT_BASE_URL,
    DEFAULT_MAIL_PAGE_SIZE,
    buildOutlookEmailHeaders,
    joinOutlookEmailUrl,
    normalizeOutlookEmailAccount,
    normalizeOutlookEmailAccounts,
    normalizeOutlookEmailAddress,
    normalizeOutlookEmailBaseUrl,
    normalizeOutlookEmailMailApiMessages,
    normalizeOutlookEmailMessage,
    normalizeOutlookEmailReceivedDateTime,
  };
});
