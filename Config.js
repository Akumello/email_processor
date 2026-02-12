/**
 * Sheet name constants. Never hardcode sheet names elsewhere.
 * @enum {string}
 */
const SHEET_NAMES = {
  RAW_EMAILS: 'Raw Emails',
  OUTPUT: 'Output'
};

/**
 * Email scanning configuration.
 * Modify these fields to filter which emails are processed.
 * @type {{cutoffDate: string, query: string, label: string, from: string, subject: string, maxResults: number}}
 */
const EMAIL_CONFIG = {
  /** Absolute cutoff — emails before this date are never processed (YYYY/MM/DD) */
  cutoffDate: '2026/02/10',
  /** Base Gmail search query (default: all inbox mail) */
  query: 'in:inbox',
  /** Optional Gmail label filter (e.g. 'clients') — leave empty to skip */
  label: '',
  /** Optional sender filter (e.g. 'someone@example.com') — leave empty to skip */
  from: '',
  /** Optional subject keyword filter — leave empty to skip */
  subject: '',
  /** Max threads to fetch per scan (Gmail caps at 500) */
  maxResults: 100
};
