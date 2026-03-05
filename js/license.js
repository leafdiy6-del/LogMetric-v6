/* ============================================================================
   license.js - 激活验证 (License & Supabase)
   依赖: config.js, core.js (无)
   被依赖: app.js
   ============================================================================ */

const LMP_SUPABASE_URL = (typeof window !== 'undefined' && window.LMP_CONFIG && window.LMP_CONFIG.SUPABASE_URL) || 'https://fqifxpxoeoyvzowiexus.supabase.co';
const LMP_ANON_KEY = (typeof window !== 'undefined' && window.LMP_CONFIG && window.LMP_CONFIG.SUPABASE_ANON_KEY) || 'sb_publishable_hP3yTSaSOACyHYd0vvbUNA_GvVDPppJ';
const LMP_LICENSE_STORAGE_KEY = 'lmp_license_key';
const LMP_DEVICE_FP_STORAGE_KEY = 'lmp_device_fp';
const LMP_LAST_VALIDATION_KEY = 'lmp_last_validation_time';
const LMP_EXPIRES_AT_KEY = 'lmp_expires_at';
const LMP_HEARTBEAT_INTERVAL_MS = 5 * 60 * 1000;
const OFFLINE_GRACE_PERIOD_DAYS = 160;
const OFFLINE_BUFFER_DAYS = 30;

let lmpHeartbeatTimerId = null;
let lmpSupabaseClient = null;

function getLicenseOverlay() { return document.getElementById('licenseOverlay'); }
function getLicenseErrorMsg() { return document.getElementById('licenseErrorMsg'); }
function getLicenseLoadingMsg() { return document.getElementById('licenseLoadingMsg'); }
function getLicenseKeyInput() { return document.getElementById('licenseKeyInput'); }
function getLicenseActivateBtn() { return document.getElementById('licenseActivateBtn'); }

function showLicenseOverlay() {
    document.documentElement.classList.remove('license-valid');
    const el = getLicenseOverlay();
    if (el) el.classList.remove('hidden');
}

function hideLicenseOverlay() {
    document.documentElement.classList.add('license-valid');
    const el = getLicenseOverlay();
    if (el) el.classList.add('hidden');
}

function clearLicenseCache() {
    localStorage.removeItem(LMP_LICENSE_STORAGE_KEY);
    localStorage.removeItem(LMP_DEVICE_FP_STORAGE_KEY);
    localStorage.removeItem(LMP_LAST_VALIDATION_KEY);
    localStorage.removeItem(LMP_EXPIRES_AT_KEY);
}

function saveCredentials(expiresAt) {
    try {
        localStorage.setItem(LMP_LAST_VALIDATION_KEY, String(Date.now()));
        if (expiresAt != null && expiresAt !== '') {
            const val = typeof expiresAt === 'string' ? expiresAt : (expiresAt && expiresAt.toISOString ? expiresAt.toISOString() : String(expiresAt));
            localStorage.setItem(LMP_EXPIRES_AT_KEY, val);
        }
    } catch (e) {}
}

function checkOfflineConditionA() {
    try {
        const raw = localStorage.getItem(LMP_LAST_VALIDATION_KEY);
        if (!raw) return false;
        const ts = parseInt(raw, 10);
        if (isNaN(ts)) return false;
        const now = Date.now();
        const days = (now - ts) / (24 * 60 * 60 * 1000);
        return days < OFFLINE_GRACE_PERIOD_DAYS;
    } catch (e) {
        return false;
    }
}

function checkOfflineConditionB() {
    try {
        const raw = localStorage.getItem(LMP_EXPIRES_AT_KEY);
        if (!raw) return true;
        const expiresMs = new Date(raw).getTime();
        if (isNaN(expiresMs)) return true;
        const bufferMs = OFFLINE_BUFFER_DAYS * 24 * 60 * 60 * 1000;
        const deadline = expiresMs + bufferMs;
        return Date.now() < deadline;
    } catch (e) {
        return true;
    }
}

function getOfflineViolationType() {
    const a = checkOfflineConditionA();
    const b = checkOfflineConditionB();
    if (!b) return 'expired';
    if (!a) return 'offline_exceeded';
    return null;
}

function showLicenseOverlayForOfflineExceeded() {
    setLicenseError('离线时间已超过 160 天，为了您的数据安全，请连接网络验证一次授权。');
    showLicenseOverlay();
}

function showLicenseOverlayForExpired() {
    setLicenseError('您的软件使用授权及离线缓冲期均已结束，请连接网络续费。');
    showLicenseOverlay();
}

function showLicenseOverlayForOnlineExpired() {
    setLicenseError('您的授权已到期，请联系Leaff d.o.o. 继续续费。');
    showLicenseOverlay();
}

function stopHeartbeat() {
    if (lmpHeartbeatTimerId) {
        clearInterval(lmpHeartbeatTimerId);
        lmpHeartbeatTimerId = null;
    }
}

function getSupabaseClient() {
    if (!lmpSupabaseClient && typeof window.supabase !== 'undefined') {
        lmpSupabaseClient = window.supabase.createClient(LMP_SUPABASE_URL, LMP_ANON_KEY);
    }
    return lmpSupabaseClient;
}

function setLicenseError(msg) {
    const el = getLicenseErrorMsg();
    if (el) el.textContent = msg || '';
}

function setLicenseLoading(show) {
    const loading = getLicenseLoadingMsg();
    const btn = getLicenseActivateBtn();
    if (loading) loading.style.display = show ? 'block' : 'none';
    if (btn) btn.disabled = !!show;
}

const LMP_ERROR_MAP = {
    invalid_license: '激活码无效，请检查后重试',
    license_revoked: '该激活码已被吊销',
    license_expired: '激活码已过期',
    device_limit_reached: '该激活码绑定的设备数量已达上限',
    device_not_bound: '当前设备未绑定，请重新激活',
    network_error: '网络错误，请稍后重试'
};

async function runHeartbeatCheck(licenseKey, deviceFp) {
    const supabase = getSupabaseClient();
    if (!supabase) return { valid: false, networkError: true };
    try {
        const { data, error } = await supabase.rpc('heartbeat_check', {
            p_license_key: licenseKey,
            p_device_fingerprint: deviceFp
        });
        if (error) return { valid: false, networkError: true };
        return Object.assign({ networkError: false }, data || {});
    } catch (e) {
        return { valid: false, networkError: true };
    }
}

async function runValidateAndBind(licenseKey, deviceFp) {
    const supabase = getSupabaseClient();
    if (!supabase) return { valid: false, message: 'network_error', networkError: true };
    try {
        const { data, error } = await supabase.rpc('validate_and_bind_license', {
            p_license_key: licenseKey,
            p_device_fingerprint: deviceFp
        });
        if (error) return { valid: false, message: 'network_error', networkError: true };
        return Object.assign({ networkError: false }, data || {});
    } catch (e) {
        return { valid: false, message: 'network_error', networkError: true };
    }
}

async function getDeviceFingerprint() {
    if (typeof window.FingerprintJS !== 'undefined') {
        const fp = await window.FingerprintJS.load();
        const result = await fp.get();
        return result.visitorId || '';
    }
    return '';
}

function startHeartbeat() {
    stopHeartbeat();
    const licenseKey = localStorage.getItem(LMP_LICENSE_STORAGE_KEY);
    const deviceFp = localStorage.getItem(LMP_DEVICE_FP_STORAGE_KEY);
    if (!licenseKey || !deviceFp) return;
    lmpHeartbeatTimerId = setInterval(async () => {
        const res = await runHeartbeatCheck(licenseKey, deviceFp);
        if (res && res.valid) {
            saveCredentials(res.expires_at);
            return;
        }
        if (res && res.networkError) {
            const violation = getOfflineViolationType();
            if (!violation) return;
            stopHeartbeat();
            if (violation === 'expired') showLicenseOverlayForExpired();
            else showLicenseOverlayForOfflineExceeded();
            return;
        }
        stopHeartbeat();
        clearLicenseCache();
        showLicenseOverlayForOnlineExpired();
    }, LMP_HEARTBEAT_INTERVAL_MS);
}

function isLocalLicenseValid() {
    const licenseKey = localStorage.getItem(LMP_LICENSE_STORAGE_KEY);
    const deviceFp = localStorage.getItem(LMP_DEVICE_FP_STORAGE_KEY);
    if (!licenseKey || !deviceFp) return false;
    return getOfflineViolationType() === null;
}

function initLicenseCheck() {
    const licenseKey = localStorage.getItem(LMP_LICENSE_STORAGE_KEY);
    const deviceFp = localStorage.getItem(LMP_DEVICE_FP_STORAGE_KEY);
    if (!licenseKey || !deviceFp) {
        showLicenseOverlay();
        return;
    }
    const violation = getOfflineViolationType();
    if (violation) {
        if (violation === 'expired') showLicenseOverlayForExpired();
        else showLicenseOverlayForOfflineExceeded();
        return;
    }
    hideLicenseOverlay();
    startHeartbeat();
    runHeartbeatCheck(licenseKey, deviceFp).then(function(res) {
        if (res && res.valid) {
            saveCredentials(res.expires_at);
            return;
        }
        if (res && res.networkError) {
            const v = getOfflineViolationType();
            if (v) {
                stopHeartbeat();
                if (v === 'expired') showLicenseOverlayForExpired();
                else showLicenseOverlayForOfflineExceeded();
            }
            return;
        }
        stopHeartbeat();
        clearLicenseCache();
        showLicenseOverlayForOnlineExpired();
    });
}

async function handleLicenseActivate() {
    const input = getLicenseKeyInput();
    const licenseKey = (input && input.value || '').trim();
    if (!licenseKey) {
        setLicenseError('请输入激活码');
        return;
    }
    setLicenseError('');
    setLicenseLoading(true);
    try {
        const deviceFp = await getDeviceFingerprint();
        if (!deviceFp) {
            setLicenseError('无法获取设备指纹，请检查网络后重试');
            setLicenseLoading(false);
            return;
        }
        const res = await runValidateAndBind(licenseKey, deviceFp);
        if (res && res.valid) {
            localStorage.setItem(LMP_LICENSE_STORAGE_KEY, licenseKey);
            localStorage.setItem(LMP_DEVICE_FP_STORAGE_KEY, deviceFp);
            saveCredentials(res.expires_at);
            hideLicenseOverlay();
            startHeartbeat();
        } else {
            const msg = res && res.message ? (LMP_ERROR_MAP[res.message] || res.message) : '激活失败，请重试';
            setLicenseError(msg);
        }
    } catch (e) {
        setLicenseError('网络错误，请稍后重试');
    }
    setLicenseLoading(false);
}
