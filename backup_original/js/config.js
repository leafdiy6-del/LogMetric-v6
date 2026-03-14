/* ============================================================================
   LogMetric Pro - 外部配置 (可部署时替换)
   ============================================================================
   注意：Supabase anon key 为 publishable，但仍建议：
   - 部署时用环境变量或构建脚本替换此文件
   - 在 Supabase 后台配置 RLS 和 rate limit 限制滥用
   ============================================================================ */
(function() {
    'use strict';
    window.LMP_CONFIG = window.LMP_CONFIG || {};
    window.LMP_CONFIG.SUPABASE_URL = 'https://fqifxpxoeoyvzowiexus.supabase.co';
    window.LMP_CONFIG.SUPABASE_ANON_KEY = 'sb_publishable_hP3yTSaSOACyHYd0vvbUNA_GvVDPppJ';
})();
