'use strict';

// Applied in memory only. The ordinary desktop configuration is never rewritten.
function applyMcpConfig(config, mode) {
  if (!['headless', 'login'].includes(mode)) return;
  config.graphApi = { ...config.graphApi, enabled: true };
  config.minimized = mode === 'headless';
  config.trayIconEnabled = false;
  config.closeAppOnCross = true;
  config.webDebug = false;
  config.disableNotifications = true;
  config.disableNotificationSound = true;
  config.disableNotificationWindowFlash = true;
  config.multiAccount = { ...config.multiAccount, enabled: false };
  config.mqtt = { ...config.mqtt, enabled: false };
  config.dbusControl = { ...config.dbusControl, enabled: false };
  config.quickChat = { ...config.quickChat, enabled: false };
}

module.exports = { applyMcpConfig };
