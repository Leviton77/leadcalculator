// ---- Edit these if any of the source addresses or wording ever changes ----

const LEAD_SENDER = 'leads@soumissionsdemenageurs.ca';
const MY_EMAIL = 'tom@tommoving.ca';
const RENEWAL_SENDER = 'admin@soumissionsdemenageurs.ca';
const RENEWAL_SUBJECT = 'Renouvellement automatique bloc de soumissions';
const RETOUR_KEYWORD = 'RETOUR';

const DEFAULT_CITY = 'Gatineau';
// Apply this Gmail label to a lead thread to mark it as an Ottawa lead
// instead of the default city (Gatineau). Useful until/unless the vendor
// starts tagging the city in the email itself.
const OTTAWA_LABEL = 'LeadsCalc/City-Ottawa';

const DEFAULT_INVOICE_TOTAL = 574.88;

const LABEL_LEAD_PROCESSED = 'LeadsCalc/Lead-Processed';
const LABEL_RETOUR_PROCESSED = 'LeadsCalc/Retour-Processed';
const LABEL_RENEWAL_PROCESSED = 'LeadsCalc/Renewal-Processed';

const SHEET_CONFIG = 'Config';
const SHEET_LEADS = 'Leads Log';
const SHEET_PERIODS = 'Renewal Periods';

const TRIGGER_INTERVAL_MINUTES = 15;
