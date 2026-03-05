/* ============================================================================
   sounds.js - 专业键盘音效 (Web Audio API 原生实现)
   风格：轻快、干脆，带实木敲击/精密仪器反馈感
   支持快速连续点击重叠播放
   ============================================================================ */

(function() {
    'use strict';

    let audioCtx = null;

    function getAudioContext() {
        if (!audioCtx) {
            audioCtx = new (window.AudioContext || window.webkitAudioContext)();
        }
        return audioCtx;
    }

    function resumeContext() {
        const ctx = getAudioContext();
        if (ctx.state === 'suspended') ctx.resume();
    }

    /**
     * 生成短促敲击音（实木/精密仪器风格）
     * @param {Object} opts - { freq, decay, type, gain }
     */
    function playTone(opts) {
        try {
            const ctx = getAudioContext();
            resumeContext();

            const freq = opts.freq || 520;
            const decay = opts.decay || 0.035;
            const type = opts.type || 'sine';
            const gainVal = opts.gain !== undefined ? opts.gain : 0.15;

            const osc = ctx.createOscillator();
            const gainNode = ctx.createGain();

            osc.connect(gainNode);
            gainNode.connect(ctx.destination);

            osc.type = type;
            osc.frequency.setValueAtTime(freq, ctx.currentTime);

            gainNode.gain.setValueAtTime(0, ctx.currentTime);
            gainNode.gain.linearRampToValueAtTime(gainVal, ctx.currentTime + 0.002);
            gainNode.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + decay);

            osc.start(ctx.currentTime);
            osc.stop(ctx.currentTime + decay);
        } catch (e) {}
    }

    /** 数字键 0-9：音调微调，模拟不同敲击感 */
    const NUM_PITCHES = { '0': 380, '1': 420, '2': 450, '3': 480, '4': 510, '5': 540, '6': 570, '7': 600, '8': 630, '9': 660 };
    const DOT_PITCH = 500;

    /**
     * 播放按键音（供外部调用，调用方负责检查 appSettings.keySound）
     * @param {string} key - '0'-'9', '.', 'del', 'ok', 'nxt', 'grade'
     */
    function playKeySound(key) {

        const k = String(key).toLowerCase();
        if (NUM_PITCHES[k] !== undefined) {
            playTone({ freq: NUM_PITCHES[k], decay: 0.03, gain: 0.12 });
        } else if (k === '.') {
            playTone({ freq: DOT_PITCH, decay: 0.025, gain: 0.1 });
        } else if (k === 'del') {
            playTone({ freq: 280, decay: 0.04, type: 'sine', gain: 0.1 });
        } else if (k === 'ok' || k === 'nxt') {
            playTone({ freq: 720, decay: 0.045, gain: 0.18 });
        } else if (k === 'grade') {
            playTone({ freq: 320, decay: 0.05, type: 'sine', gain: 0.07 });
        }
    }

    window.playKeySound = playKeySound;

    document.addEventListener('click', function() { resumeContext(); }, { once: true });
})();
