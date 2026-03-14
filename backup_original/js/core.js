/* ============================================================================
   core.js - 常量与工具函数 (Constants & Utilities)
   依赖: config.js (可选)
   被依赖: i18n.js, license.js, app.js
   ============================================================================ */

/* ----------------------------------------------------------------------------
   LocalStorage 键名配置 (Storage Keys)
   ---------------------------------------------------------------------------- */
const STORAGE_KEY = 'oak_v286_data';
const HIST_KEY = 'oak_v286_hist';
const LANG_KEY = 'oak_v286_lang';
const QUICK_KEY = 'oak_v286_quick';
const SETTINGS_KEY = 'oak_v286_settings';
const MIX_STATE_KEY = 'oak_v286_mix';
const SNAPSHOTS_KEY = 'oak_v286_snapshots';
const SESSION_KEY = 'oak_v286_session';

/* ----------------------------------------------------------------------------
   业务常量 (Business Constants)
   ---------------------------------------------------------------------------- */
const GRADES = ['F', 'A+', 'A', 'B', 'C', 'D'];
const COMPANY_FIELDS = ['name', 'address', 'city', 'zip', 'phone', 'email', 'website', 'taxId', 'bank'];

function defaultCompany() {
    return { name: '', address: '', city: '', zip: '', phone: '', email: '', website: '', taxId: '', bank: '' };
}
function defaultSeller() {
    return { name: '', address: '', city: '', zip: '', phone: '', email: '', website: '', taxId: '', bank: '', type: 'buyer' };
}
function normalizeCompany(v) {
    if (v == null) return defaultCompany();
    if (typeof v === 'string') return Object.assign(defaultCompany(), { name: v });
    const o = defaultCompany();
    COMPANY_FIELDS.forEach(f => { if (v[f] != null) o[f] = String(v[f]); });
    return o;
}
function normalizeSeller(v) {
    if (v == null) return defaultSeller();
    if (typeof v === 'string') return Object.assign(defaultSeller(), { name: v });
    const o = defaultSeller();
    COMPANY_FIELDS.forEach(f => { if (v[f] != null) o[f] = String(v[f]); });
    if (v.type === 'seller' || v.type === 'buyer') o.type = v.type;
    return o;
}
function getCompanyName(obj) {
    if (obj == null) return '';
    if (typeof obj === 'string') return obj;
    return (obj.name != null && obj.name !== '') ? String(obj.name) : '';
}
function migrateHistoriesSeller(h) {
    if(!h || !h.seller) return;
    h.seller = (h.seller || []).map(s => typeof s === 'string' ? Object.assign(defaultSeller(), { name: s }) : normalizeSeller(s));
}
function migrateHistoriesCompany(h) {
    if(!h || !h.company) return;
    h.company = (h.company || []).map(c => typeof c === 'string' ? Object.assign(defaultCompany(), { name: c }) : normalizeCompany(c));
}
