const fs = require('fs');
const path = require('path');
const React = require('react');
const { shell } = require('electron');
const vortexApi = require('vortex-api');
const pendingLLInstalls = new Map();

const SUPPORTED_GAME_IDS = new Set(['skyrim', 'skyrimse', 'skyrimvr']);
let snapshotTimer = null;
let commandPollerStarted = false;

function readJson(filePath) {
  try {
    return JSON.parse(fs.readFileSync(filePath, 'utf8'));
  } catch (err) {
    return {};
  }
}

function localAppData() {
  return process.env.LOCALAPPDATA || path.join(process.env.USERPROFILE || '', 'AppData', 'Local');
}

function extensionConfigPath() {
  return path.join(__dirname, 'll-integration.config.json');
}

function nativeConfigPath() {
  const extensionConfig = readJson(extensionConfigPath());
  return extensionConfig.nativeConfigPath
    || path.join(localAppData(), 'LLIntegration', 'native-app', 'config.json');
}

function nativeAppPath() {
  const extensionConfig = readJson(extensionConfigPath());
  return extensionConfig.nativeAppPath
    || path.join(localAppData(), 'LLIntegration', 'native-app');
}

function extensionLogPath() {
  return path.join(__dirname, 'll-integration.log');
}

function nativeLogPath() {
  return path.join(nativeAppPath(), 'll-integration-extension.log');
}

function formatError(err) {
  if (!err) {
    return '';
  }
  return err.stack || err.message || String(err);
}

function cleanDownloadDisplayName(value) {
  return String(value || '')
    .replace(/^\s*(download\s+your\s+files?|download\s+files?|files?)\s*[-:–—]\s*/i, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function logLine(message, payload) {
  const suffix = payload === undefined ? '' : ` ${typeof payload === 'string' ? payload : JSON.stringify(payload)}`;
  const line = `${new Date().toISOString()} ${message}${suffix}\n`;
  [extensionLogPath(), nativeLogPath()].forEach((filePath) => {
    try {
      fs.mkdirSync(path.dirname(filePath), { recursive: true });
      fs.appendFileSync(filePath, line, 'utf8');
    } catch (err) {
      // Best-effort debug log only.
    }
  });
}

function writeJson(filePath, data) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  const tmpPath = `${filePath}.tmp`;
  fs.writeFileSync(tmpPath, JSON.stringify(data, null, 2), 'utf8');
  fs.renameSync(tmpPath, filePath);
}

function safeGet(obj, pathParts, fallback) {
  return pathParts.reduce(
    (value, key) => (value !== undefined && value !== null ? value[key] : undefined),
    obj,
  ) ?? fallback;
}

function displayModName(mod) {
  return safeGet(mod, ['attributes', 'customFileName'])
    || safeGet(mod, ['attributes', 'name'])
    || safeGet(mod, ['attributes', 'logicalFileName'])
    || mod.id;
}

function vortexStatePath() {
  return path.join(nativeAppPath(), 'vortex_state.json');
}

function vortexCommandPath() {
  return path.join(nativeAppPath(), 'vortex_commands.json');
}

function vortexCommandResultsPath() {
  return path.join(nativeAppPath(), 'vortex_command_results.json');
}

function currentVortexSnapshot(api) {
  const state = api.store.getState();
  let activeGameId = '';
  let profile = {};
  try {
    activeGameId = vortexApi.selectors.activeGameId(state) || '';
  } catch (err) {
    logLine('activeGameId selector failed', formatError(err));
  }
  try {
    profile = vortexApi.selectors.activeProfile(state) || {};
  } catch (err) {
    logLine('activeProfile selector failed', formatError(err));
  }
  const gameId = profile.gameId || activeGameId;
  const mods = safeGet(state, ['persistent', 'mods', gameId], {});
  const downloads = safeGet(state, ['persistent', 'downloads', 'files'], {});
  let stagingPath = '';
  let downloadsPath = '';
  try {
    stagingPath = vortexApi.selectors.installPathForGame(state, gameId) || '';
  } catch (err) {
    try {
      stagingPath = vortexApi.selectors.installPath(state) || '';
    } catch (fallbackErr) {
      logLine('staging path selector failed', formatError(fallbackErr));
    }
  }
  try {
    downloadsPath = vortexApi.selectors.downloadPathForGame(state, gameId) || '';
  } catch (err) {
    try {
      downloadsPath = vortexApi.selectors.downloadPath(state) || '';
    } catch (fallbackErr) {
      logLine('download path selector failed', formatError(fallbackErr));
    }
  }

  return {
    capturedAt: new Date().toISOString(),
    activeGameId: gameId || '',
    activeProfileId: profile.id || '',
    activeProfileName: profile.name || '',
    stagingPath,
    downloadsPath,
    mods: Object.entries(mods).map(([id, mod]) => ({
      id,
      name: displayModName({ id, ...mod }),
      state: mod.state || '',
      enabled: !!safeGet(profile, ['modState', id, 'enabled'], false),
      installationPath: mod.installationPath || id,
      archiveId: mod.archiveId || '',
      version: safeGet(mod, ['attributes', 'version'], ''),
      source: safeGet(mod, ['attributes', 'source'], ''),
      modId: safeGet(mod, ['attributes', 'modId'], ''),
      fileId: safeGet(mod, ['attributes', 'fileId'], ''),
    })),
    downloads: Object.entries(downloads).map(([id, download]) => ({
      id,
      game: download.game || '',
      state: download.state || '',
      localPath: download.localPath || '',
      fileName: path.basename(download.localPath || download.fileName || id),
      modInfo: download.modInfo || {},
    })),
  };
}

function parseLLIni(filePath) {
  const result = {};
  if (!filePath || !fs.existsSync(filePath)) {
    return result;
  }

  const lines = fs.readFileSync(filePath, 'utf8').split(/\r?\n/);
  let section = '';

  lines.forEach((rawLine) => {
    const line = rawLine.trim();
    if (!line || line.startsWith('#') || line.startsWith(';')) {
      return;
    }

    const sectionMatch = line.match(/^\[(.+)]$/);
    if (sectionMatch) {
      section = sectionMatch[1].trim();
      return;
    }

    const eq = line.indexOf('=');
    if (eq < 0) {
      return;
    }

    if (section && section !== 'LoversLab') {
      return;
    }

    const key = line.slice(0, eq).trim();
    const value = line.slice(eq + 1).trim();
    result[key] = value;
  });

  return result;
}

function writeVortexSnapshot(api) {
  const snapshot = currentVortexSnapshot(api);
  try {
    writeJson(vortexStatePath(), snapshot);
  } catch (err) {
    logLine('snapshot write failed', formatError(err));
  }

  const configPath = nativeConfigPath();
  const config = readJson(configPath);
  config.vortex_state_path = vortexStatePath();
  if (snapshot.downloadsPath) {
    config.vortex_downloads_path = snapshot.downloadsPath;
  }
  if (snapshot.stagingPath) {
    config.vortex_staging_path = snapshot.stagingPath;
    config.vortex_mods_path = snapshot.stagingPath;
  }
  if (snapshot.activeGameId) {
    config.active_vortex_game = snapshot.activeGameId;
  }
  config.active_vortex_profile_id = snapshot.activeProfileId || '';
  config.active_vortex_profile = snapshot.activeProfileName || snapshot.activeProfileId || '';
  config.active_vortex_synced_at = snapshot.capturedAt;
  try {
    writeJson(configPath, config);
  } catch (err) {
    logLine('config sync failed', formatError(err));
  }
}

function readCommandQueue() {
  const data = readJson(vortexCommandPath());
  return Array.isArray(data.commands) ? data.commands : [];
}

function writeCommandQueue(commands) {
  writeJson(vortexCommandPath(), { commands });
}

function appendCommandResult(result) {
  const existing = readJson(vortexCommandResultsPath());
  const results = Array.isArray(existing.results) ? existing.results : [];
  results.push({ ...result, completedAt: new Date().toISOString() });
  writeJson(vortexCommandResultsPath(), { results: results.slice(-100) });
}

function commandId() {
  return `ll-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function applyDownloadMetadata(api, downloadId, gameId, archivePath, command = {}) {
  const archiveName = path.basename(archivePath);
  const sourceUrl = String(command.sourceUrl || command.pageUrl || command.modHomepage || '');
  const downloadUrl = String(command.downloadUrl || command.download_url || '');
  const sourceType = String(command.sourceType || command.source_type || (
    sourceUrl.toLowerCase().includes('loverslab.com') || downloadUrl.toLowerCase().includes('loverslab.com')
      ? 'loverslab'
      : 'nexus'
  )).toLowerCase();

  const isLL = sourceType === 'loverslab' || sourceUrl.toLowerCase().includes('loverslab.com');
  const vortexSource = sourceUrl ? 'website' : (isLL ? 'website' : (command.sourceName || 'Nexus Mods'));
  const displayName =
    cleanDownloadDisplayName(command.downloadName)
    || cleanDownloadDisplayName(path.basename(archiveName, path.extname(archiveName)))
    || path.basename(archiveName, path.extname(archiveName));
  const version = command.version || '';

  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'source', vortexSource));
  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'sourceType', isLL ? 'loverslab' : sourceType));
  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'llSourceName', isLL ? 'LoversLab' : (command.displaySource || command.sourceName || sourceType)));

  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'game', gameId));
  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'name', displayName));
  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'version', version));
  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'homepage', sourceUrl));
  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'website', sourceUrl));
  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'url', sourceUrl));
  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'fileName', archiveName));
  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'logicalFileName', archiveName));

  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'llIntegration', true));
  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'llIntegrationKind', command.operation || ''));
  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'llVoicePack', command.voicePack === true));
  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'llVoiceForBaseMod', command.voiceForBaseMod || ''));
  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'llVoiceForBaseInternalName', command.voiceForBaseInternalName || ''));
  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'llVoiceCategory', command.voiceCategory || ''));

  if (command.nexusFileId) {
    api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'fileId', command.nexusFileId));
  }
}

function ensureDownloadRecord(api, gameId, archivePath, command = {}) {
  const state = api.store.getState();
  const downloads = safeGet(state, ['persistent', 'downloads', 'files'], {});
  const archiveName = path.basename(archivePath);

  const existing = Object.entries(downloads).find(([, download]) => (
    path.basename(download.localPath || download.fileName || '').toLowerCase() === archiveName.toLowerCase()
  ));

  if (existing) {
    applyDownloadMetadata(api, existing[0], gameId, archivePath, command);
    return existing[0];
  }

  const downloadsPath = vortexApi.selectors.downloadPathForGame(state, gameId);
  const targetPath = path.join(downloadsPath, archiveName);

  if (path.resolve(archivePath).toLowerCase() !== path.resolve(targetPath).toLowerCase()) {
    fs.copyFileSync(archivePath, targetPath);
  }

  const stats = fs.statSync(targetPath);
  const id = commandId();

  api.store.dispatch(vortexApi.actions.addLocalDownload(id, gameId, archiveName, stats.size));
  applyDownloadMetadata(api, id, gameId, targetPath, command);

  return id;
}

function removeFileIfExists(filePath) {
  if (!filePath) {
    return;
  }
  try {
    if (fs.existsSync(filePath)) {
      fs.unlinkSync(filePath);
    }
  } catch (err) {
    logLine('file delete failed', { filePath, error: formatError(err) });
  }
}

function safePathExists(targetPath) {
  try {
    return !!targetPath && fs.existsSync(targetPath);
  } catch (err) {
    return false;
  }
}

function safeRmDir(targetPath) {
  if (!safePathExists(targetPath)) {
    return;
  }
  try {
    fs.rmSync(targetPath, { recursive: true, force: true });
  } catch (err) {
    logLine('directory delete failed', { targetPath, error: formatError(err) });
    throw err;
  }
}

function copyDirRecursive(source, target) {
  if (!safePathExists(source)) {
    return;
  }
  fs.mkdirSync(target, { recursive: true });
  for (const entry of fs.readdirSync(source, { withFileTypes: true })) {
    const sourcePath = path.join(source, entry.name);
    const targetPath = path.join(target, entry.name);
    if (entry.isDirectory()) {
      copyDirRecursive(sourcePath, targetPath);
    } else if (entry.isFile()) {
      fs.copyFileSync(sourcePath, targetPath);
    }
  }
}

function backupRootPath() {
  return path.join(nativeAppPath(), 'backups', 'vortex-updates');
}

function timestampForPath() {
  return new Date().toISOString().replace(/[:.]/g, '-');
}

function backupOldModFolder(command) {
  const oldFolder = command.oldInstalledFolder || '';
  if (command.backupOldModFolder !== true || !safePathExists(oldFolder)) {
    return '';
  }

  const safeName = path.basename(oldFolder).replace(/[<>:"/\\|?*]/g, '_');
  const target = path.join(
    backupRootPath(),
    `${timestampForPath()}-${safeName || command.replaceModId || 'old-mod'}`,
  );

  try {
    copyDirRecursive(oldFolder, target);
    logLine('old mod folder backed up', { oldFolder, target });
    return target;
  } catch (err) {
    logLine('old mod folder backup failed', { oldFolder, target, error: formatError(err) });
    throw err;
  }
}

function findOldInstallFolderFromSnapshot(api, command, gameId) {
  if (command.oldInstalledFolder && safePathExists(command.oldInstalledFolder)) {
    return command.oldInstalledFolder;
  }

  const state = api.store.getState();
  const mods = safeGet(state, ['persistent', 'mods', gameId], {});
  const oldMod = mods[command.replaceModId || ''];
  if (!oldMod) {
    return '';
  }

  let stagingPath = '';
  try {
    stagingPath = vortexApi.selectors.installPathForGame(state, gameId) || '';
  } catch (err) {
    try {
      stagingPath = vortexApi.selectors.installPath(state) || '';
    } catch (fallbackErr) {
      stagingPath = '';
    }
  }

  const installRel = oldMod.installationPath || command.replaceModId || '';
  const candidate = stagingPath && installRel ? path.join(stagingPath, installRel) : '';
  return safePathExists(candidate) ? candidate : '';
}

async function removeVortexModEntry(api, gameId, modId) {
  if (!modId) {
    return;
  }

  const state = api.store.getState();
  const mods = safeGet(state, ['persistent', 'mods', gameId], {});
  if (mods[modId] === undefined) {
    return;
  }

  await new Promise((resolve, reject) => {
    api.events.emit('remove-mod', gameId, modId, (err) => {
      if (err) {
        reject(err);
      } else {
        resolve();
      }
    }, {
      silent: true,
      incomplete: true,
      ignoreInstalling: true,
    });
  });
}

async function cleanupOldReplacement(api, command, gameId, phase) {
  const oldModId = command.replaceModId || '';
  if (!oldModId) {
    return { backupPath: '' };
  }

  const oldFolder = findOldInstallFolderFromSnapshot(api, command, gameId);
  let backupPath = '';

  if (phase === 'before') {
    if (oldFolder) {
      backupPath = backupOldModFolder({ ...command, oldInstalledFolder: oldFolder });
    }

    try {
      await removeVortexModEntry(api, gameId, oldModId);
    } catch (err) {
      logLine('old Vortex mod entry remove failed', {
        modId: oldModId,
        phase,
        error: formatError(err),
      });
    }

    if (command.deleteOldModFolder === true && oldFolder && safePathExists(oldFolder)) {
      safeRmDir(oldFolder);
      logLine('old mod folder deleted', { oldFolder });
    }
  }

  if (phase === 'after' && command.cleanupDuplicateAfterInstall === true) {
    try {
      await removeVortexModEntry(api, gameId, oldModId);
    } catch (err) {
      logLine('post-install duplicate remove failed', {
        modId: oldModId,
        phase,
        error: formatError(err),
      });
    }

    if (oldFolder && safePathExists(oldFolder) && command.deleteOldModFolder === true) {
      safeRmDir(oldFolder);
      logLine('old mod folder deleted after install', { oldFolder });
    }
  }

  return { backupPath };
}

async function removeOldVortexMod(api, command, gameId) {
  const oldModId = command.replaceModId || '';
  if (!oldModId || command.removeOldBeforeInstall !== true) {
    return { backupPath: '' };
  }

  const result = await cleanupOldReplacement(api, command, gameId, 'before');

  if (command.oldDownloadId) {
    try {
      await new Promise((resolve, reject) => {
        api.events.emit('remove-download', command.oldDownloadId, (err) => {
          if (err) {
            reject(err);
          } else {
            resolve();
          }
        }, {
          silent: true,
          confirmed: true,
        });
      });
    } catch (err) {
      logLine('old download remove failed', {
        downloadId: command.oldDownloadId,
        error: formatError(err),
      });
    }
  }

  if (command.deleteOldArchive === true) {
    removeFileIfExists(command.oldArchivePath || '');
    removeFileIfExists(command.oldSidecarPath || '');
    removeFileIfExists(command.oldMetadataSidecarPath || '');
  }

  writeVortexSnapshot(api);
  return result;
}

function llModAttributesFromCommand(command, archivePath) {
  const archiveName = path.basename(archivePath || command.archivePath || command.archiveName || '');
  const sourceUrl = String(command.sourceUrl || command.pageUrl || command.modHomepage || '');
  const downloadUrl = String(command.downloadUrl || command.download_url || '');
  const sourceType = String(command.sourceType || command.source_type || (
    sourceUrl.toLowerCase().includes('loverslab.com') || downloadUrl.toLowerCase().includes('loverslab.com')
      ? 'loverslab'
      : 'nexus'
  )).toLowerCase();

  const isLL = sourceType === 'loverslab' || sourceUrl.toLowerCase().includes('loverslab.com');

  return {
    // IMPORTANT: Vortex Source enum does not accept "LoversLab" as native source.
    source: sourceUrl ? 'website' : (isLL ? 'website' : (command.sourceName || 'nexus')),
    sourceType: isLL ? 'loverslab' : sourceType,

    website: sourceUrl,
    homepage: sourceUrl,
    url: sourceUrl,

    version: command.version || '',
    fileName: archiveName,
    logicalFileName: archiveName,

    llIntegration: true,
    llIntegrationKind: command.operation || '',
    llSourceName: isLL ? 'LoversLab' : (command.displaySource || command.sourceName || sourceType),
    llPageUrl: sourceUrl,
    llDownloadUrl: downloadUrl,

    llVoicePack: command.voicePack === true,
    llVoiceForBaseMod: command.voiceForBaseMod || '',
    llVoiceForBaseInternalName: command.voiceForBaseInternalName || '',
    llVoiceCategory: command.voiceCategory || '',
    llDownloadName: cleanDownloadDisplayName(command.downloadName) || archiveName,
  };
}

function applyInstalledModMetadata(api, gameId, modId, archivePath, command) {
  if (!modId) {
    logLine('applyInstalledModMetadata skipped: missing modId', { gameId, archivePath });
    return;
  }

  const attrs = llModAttributesFromCommand(command, archivePath);
  const actions = vortexApi.actions;

  logLine('applying installed mod metadata', { gameId, modId, attrs });

  // On log les actions disponibles pour confirmer le bon nom dans ta version Vortex.
  logLine('available metadata actions', Object.keys(actions).filter((key) => (
    /mod|attribute|info|meta/i.test(key)
  )));

  try {
    if (typeof actions.setModAttribute === 'function') {
      Object.entries(attrs).forEach(([key, value]) => {
        actions.setModAttribute.length >= 4
          ? api.store.dispatch(actions.setModAttribute(gameId, modId, key, value))
          : api.store.dispatch(actions.setModAttribute(modId, key, value));
      });
      return;
    }

    if (typeof actions.setModAttributes === 'function') {
      actions.setModAttributes.length >= 3
        ? api.store.dispatch(actions.setModAttributes(gameId, modId, attrs))
        : api.store.dispatch(actions.setModAttributes(modId, attrs));
      return;
    }

    logLine('no installed mod metadata action found', Object.keys(actions));
  } catch (err) {
    logLine('applyInstalledModMetadata failed', formatError(err));
  }
}

async function handleCommand(api, command) {
  if (command.action === 'enable_mod') {
    const profileId = command.profileId || safeGet(currentVortexSnapshot(api), ['activeProfileId'], '');
    const modId = command.modId || '';
    if (!profileId || !modId) {
      return { ok: false, error: 'Missing profileId or modId for enable_mod' };
    }
    api.store.dispatch(vortexApi.actions.setModEnabled(profileId, modId, true));
    writeVortexSnapshot(api);
    return { ok: true, profileId, modId };
  }

  if (command.action !== 'install_archive') {
    return { ok: false, error: `Unknown command: ${command.action}` };
  }
  const archivePath = command.archivePath || '';
  if (!archivePath || !fs.existsSync(archivePath)) {
    return { ok: false, error: `Archive not found: ${archivePath}` };
  }
  const state = api.store.getState();
  const gameId = command.gameId || vortexApi.selectors.activeGameId(state);
  const replacementCleanup = await removeOldVortexMod(api, command, gameId);
  const downloadId = ensureDownloadRecord(api, gameId, archivePath, command);

  pendingLLInstalls.set(String(downloadId), {
    gameId,
    archivePath,
    command: { ...command },
    createdAt: Date.now(),
  });
  logLine('pending LL install registered', {
    downloadId,
    gameId,
    archivePath,
    operation: command.operation || '',
  });
  return new Promise((resolve) => {
    api.events.emit(
      'start-install-download',
      downloadId,
      { allowAutoEnable: command.allowAutoEnable !== false },
     async (err, modId) => {
      logLine('install result', {
        error: err ? formatError(err) : '',
        modId,
        downloadId,
        archivePath,
        operation: command.operation || '',
      });

      if (!err) {
        try {
          const snapshot = currentVortexSnapshot(api);
          const profileId = command.profileId || snapshot.activeProfileId || '';
          const shouldEnable = command.enableAfterInstall === true || command.oldWasEnabled === true || command.allowAutoEnable !== false;

          if (profileId && modId && shouldEnable) {
            api.store.dispatch(vortexApi.actions.setModEnabled(profileId, modId, true));
          }

          const hasCommandUrl =
            !!(command.sourceUrl || command.pageUrl || command.modHomepage);

          if (hasCommandUrl) {
            applyInstalledModMetadata(api, gameId, modId, archivePath, command);
          } else {
            logLine('post-install direct metadata skipped: metadata is handled by did-install-mod', {
              modId,
              archivePath,
              operation: command.operation || '',
            });
          }
          writeVortexSnapshot(api);

          if (command.cleanupDuplicateAfterInstall === true && command.replaceModId && command.replaceModId !== modId) {
            await cleanupOldReplacement(api, command, gameId, 'after');
          }
          
        } catch (snapshotErr) {
          logLine('post-install replacement cleanup failed', formatError(snapshotErr));
        }
      }

      resolve(err
        ? {
          ok: false,
          error: String(err.message || err),
          downloadId,
          backupPath: replacementCleanup.backupPath || '',
        }
        : {
          ok: true,
          downloadId,
          modId,
          replacedModId: command.replaceModId || '',
          backupPath: replacementCleanup.backupPath || '',
        });
      },
    );
  });
}

function parseSimpleIni(filePath) {
  const result = {};
  if (!filePath || !fs.existsSync(filePath)) {
    return result;
  }

  const lines = fs.readFileSync(filePath, 'utf8').split(/\r?\n/);
  let section = '';

  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line || line.startsWith(';') || line.startsWith('#')) {
      continue;
    }

    const sectionMatch = line.match(/^\[(.+)]$/);
    if (sectionMatch) {
      section = sectionMatch[1];
      continue;
    }

    const eq = line.indexOf('=');
    if (eq < 0) {
      continue;
    }

    const key = line.slice(0, eq).trim();
    const value = line.slice(eq + 1).trim();

    if (!section || section === 'LoversLab') {
      result[key] = value;
    }
  }

  return result;
}

function findLLSidecarForArchive(config, archivePathOrName) {
  const value = String(archivePathOrName || '');
  const archiveName = path.basename(value);
  const candidates = [];

  if (!archiveName) {
    return '';
  }

  // Case 1: full archive path was provided.
  if (value.includes('\\') || value.includes('/')) {
    candidates.push(`${value}.ll.ini`);
  }

  // Case 2: Vortex/native-app downloads path + archive name.
  if (config.vortex_downloads_path) {
    candidates.push(path.join(config.vortex_downloads_path, `${archiveName}.ll.ini`));
  }

  // Case 3: LL Integration central metadata folder.
  if (config.metadata_path) {
    candidates.push(path.join(config.metadata_path, 'downloads', `${archiveName}.ll.ini`));
  }

  // Case 4: fallback relative path, useful only for debug/dev.
  candidates.push(`${archiveName}.ll.ini`);

  const found = candidates.find((candidate) => fs.existsSync(candidate)) || '';

  if (!found) {
    logLine('LL sidecar not found', {
      archivePathOrName: value,
      archiveName,
      candidates,
    });
  } else {
    logLine('LL sidecar found', {
      archiveName,
      sidecar: found,
    });
  }

  return found;
}

function applyLLMetadataToInstalledMod(api, gameId, modId, archivePath, metadata) {
  if (!modId) {
    return false;
  }

  const archiveName = path.basename(archivePath || metadata.archive_name || metadata.file_name || '');
  const source = (metadata.source || 'loverslab').toLowerCase();

  const pageUrl =
    metadata.page_url
    || metadata.mod_homepage
    || metadata.homepage
    || metadata.website
    || metadata.url
    || '';

  const downloadUrl = metadata.download_url || '';

  const isLL = source === 'loverslab' || pageUrl.toLowerCase().includes('loverslab.com');

  const attrs = {
    source: pageUrl ? 'website' : (isLL ? 'website' : source),
    sourceType: source,

    version: metadata.version || '',

    // Vortex UI fields
    website: pageUrl,
    homepage: pageUrl,
    url: pageUrl,

    logicalFileName: archiveName,
    fileName: archiveName,

    llIntegration: true,
    llIntegrationKind: metadata.ll_integration_kind || '',
    llSourceName: isLL ? 'LoversLab' : metadata.display_source || source,
    llPageUrl: pageUrl,
    llDownloadUrl: downloadUrl,

    llVoicePack: metadata.ll_integration_kind === 'voice_pack',
    llVoiceForBaseMod: metadata.voice_for_base_mod || '',
    llVoiceForBaseInternalName: metadata.voice_for_base_internal_name || '',
    llVoiceCategory: metadata.voice_category || '',
  };

  const actions = vortexApi.actions;

  logLine('apply LL metadata to installed mod start', {
    gameId,
    modId,
    archiveName,
    pageUrl,
    attrs,
    hasSetModAttribute: typeof actions.setModAttribute === 'function',
    hasSetModAttributes: typeof actions.setModAttributes === 'function',
  });

  try {
    if (typeof actions.setModAttribute === 'function') {
      Object.entries(attrs).forEach(([key, value]) => {
        api.store.dispatch(actions.setModAttribute(gameId, modId, key, value));
      });

      logLine('apply LL metadata to installed mod done', {
        gameId,
        modId,
        archiveName,
        pageUrl,
      });

      return true;
    }

    if (typeof actions.setModAttributes === 'function') {
      api.store.dispatch(actions.setModAttributes(gameId, modId, attrs));

      logLine('apply LL metadata to installed mod done via setModAttributes', {
        gameId,
        modId,
        archiveName,
        pageUrl,
      });

      return true;
    }

    logLine('No installed-mod metadata action found', Object.keys(actions).filter((key) => (
      /mod|attribute|meta|info/i.test(key)
    )));
  } catch (err) {
    logLine('applyLLMetadataToInstalledMod failed', formatError(err));
  }

  return false;
}

function archivePathFromDownload(download) {
  return download.localPath || download.fileName || download.id || '';
}

function downloadArchivePathFromId(api, downloadId) {
  const state = api.store.getState();
  const downloads = safeGet(state, ['persistent', 'downloads', 'files'], {});
  const download = downloads[String(downloadId || '')];

  if (!download) {
    logLine('downloadArchivePathFromId no download found', {
      downloadId,
      knownDownloadIds: Object.keys(downloads).slice(0, 20),
    });
    return '';
  }

  const archivePath = archivePathFromDownload(download);
  logLine('downloadArchivePathFromId resolved', {
    downloadId,
    archivePath,
    localPath: download.localPath || '',
    fileName: download.fileName || '',
    logicalFileName: safeGet(download, ['modInfo', 'logicalFileName'], ''),
  });

  return archivePath;
}

function syncLLMetadata(api) {
  const config = readJson(nativeConfigPath());
  const state = api.store.getState();

  let gameId = '';
  try {
    gameId = vortexApi.selectors.activeGameId(state) || config.active_vortex_game || '';
  } catch (err) {
    gameId = config.active_vortex_game || '';
  }

  if (!gameId) {
    return 0;
  }

  const downloads = safeGet(state, ['persistent', 'downloads', 'files'], {});
  const mods = safeGet(state, ['persistent', 'mods', gameId], {});
  let updated = 0;

  const metadataByDownloadId = {};
  const metadataByArchiveName = {};

  Object.entries(downloads).forEach(([downloadId, download]) => {
    const archivePath = archivePathFromDownload(download);
    const archiveName = path.basename(archivePath || '');

    if (!archiveName.match(/\.(7z|zip|rar)$/i)) {
      return;
    }

    const sidecar = findLLSidecarForArchive(config, archivePath);
    if (!sidecar) {
      return;
    }

    const metadata = parseLLIni(sidecar);
    if (!metadata.page_url && !metadata.download_url) {
      return;
    }

    applyLLMetadataToDownload(api, downloadId, gameId, archivePath, metadata);

    metadataByDownloadId[downloadId] = { metadata, archivePath };
    metadataByArchiveName[archiveName.toLowerCase()] = { metadata, archivePath };

    updated += 1;
  });

  Object.entries(mods).forEach(([modId, mod]) => {
    const archiveId = mod.archiveId || '';
    const logicalFileName = safeGet(mod, ['attributes', 'logicalFileName'], '');
    const fileName = safeGet(mod, ['attributes', 'fileName'], '');

    let match =
      metadataByDownloadId[archiveId]
      || metadataByArchiveName[path.basename(logicalFileName || '').toLowerCase()]
      || metadataByArchiveName[path.basename(fileName || '').toLowerCase()];

    // Important: fallback direct. If download metadata sync failed,
    // still try to find the sidecar from the installed mod archive name.
    if (!match) {
      const archiveName =
        path.basename(logicalFileName || '')
        || path.basename(fileName || '')
        || path.basename(archiveId || '');

      if (archiveName) {
        const sidecar = findLLSidecarForArchive(config, archiveName);
        if (sidecar) {
          const metadata = parseLLIni(sidecar);
          if (metadata.page_url || metadata.download_url) {
            match = {
              metadata,
              archivePath: archiveName,
            };
          }
        }
      }
    }

    if (!match) {
      logLine('No LL metadata match for installed mod', {
        modId,
        archiveId,
        logicalFileName,
        fileName,
      });
      return;
    }

    logLine('LL metadata matched installed mod', {
      modId,
      archiveId,
      logicalFileName,
      fileName,
      sourceArchive: match.archivePath,
      pageUrl: match.metadata.page_url || '',
      version: match.metadata.version || '',
    });

    if (applyLLMetadataToInstalledMod(api, gameId, modId, match.archivePath, match.metadata)) {
      updated += 1;
    }
  });

  if (updated > 0) {
    logLine('LL metadata sync applied', { updated });
    writeVortexSnapshot(api);
  }

  return updated;
}

function applyLLMetadataToDownload(api, downloadId, gameId, archivePath, metadata) {
  const archiveName = path.basename(archivePath || metadata.archive_name || metadata.file_name || downloadId);
  const source = (metadata.source || 'loverslab').toLowerCase();

  const pageUrl =
    metadata.page_url
    || metadata.mod_homepage
    || metadata.homepage
    || metadata.website
    || metadata.url
    || '';

  const isLL = source === 'loverslab' || pageUrl.toLowerCase().includes('loverslab.com');
  const vortexSource = pageUrl ? 'website' : (isLL ? 'website' : source);

  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'source', vortexSource));
  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'sourceType', source));
  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'llSourceName', isLL ? 'LoversLab' : metadata.display_source || source));
  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'game', gameId));
  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'name', metadata.page_title || metadata.file_name || path.basename(archiveName, path.extname(archiveName))));
  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'version', metadata.version || ''));

  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'homepage', pageUrl));
  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'website', pageUrl));
  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'url', pageUrl));

  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'fileName', archiveName));
  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'logicalFileName', archiveName));
  api.store.dispatch(vortexApi.actions.setDownloadModInfo(downloadId, 'llIntegration', true));
}

function syncLLDownloadMetadata(api) {
  const config = readJson(nativeConfigPath());
  const state = api.store.getState();
  const gameId = config.active_vortex_game || vortexApi.selectors.activeGameId(state);
  const downloads = safeGet(state, ['persistent', 'downloads', 'files'], {});

  let updated = 0;

  Object.entries(downloads).forEach(([downloadId, download]) => {
    const archivePath = download.localPath || download.fileName || '';
    const archiveName = path.basename(archivePath || '');

    if (!archiveName.match(/\.(7z|zip|rar)$/i)) {
      return;
    }

    const sidecar = findLLSidecarForArchive(config, archivePath);
    if (!sidecar) {
      return;
    }

    const metadata = parseSimpleIni(sidecar);
    if (!metadata.page_url && !metadata.download_url) {
      return;
    }

    applyLLMetadataToDownload(api, downloadId, gameId, archivePath, metadata);
    updated += 1;
  });

  if (updated > 0) {
    logLine('synced LL download metadata', { updated });
    writeVortexSnapshot(api);
  }

  return updated;
}

function startCommandPoller(api) {
  if (commandPollerStarted) {
    return;
  }
  commandPollerStarted = true;
  setInterval(() => {
    try {
      const commands = readCommandQueue();
      const pending = commands.filter((command) => !command.processedAt && !command.startedAt);
      if (pending.length === 0) {
        return;
      }
      pending.forEach((command) => {
        command.startedAt = new Date().toISOString();
      });
      try {
        writeCommandQueue(commands);
      } catch (err) {
        logLine('command start write failed', formatError(err));
      }
      Promise.all(pending.map((command) => (
        handleCommand(api, command)
          .then((result) => appendCommandResult({ commandId: command.id, ...result }))
          .catch((err) => {
            logLine('command failed', formatError(err));
            appendCommandResult({ commandId: command.id, ok: false, error: String(err.message || err) });
          })
          .then(() => {
            command.processedAt = new Date().toISOString();
          })
      ))).then(() => writeCommandQueue(commands)).catch((err) => {
        logLine('command queue write failed', formatError(err));
      });
    } catch (err) {
      logLine('command poll failed', formatError(err));
    }
  }, 2000);
}

function showMessage(api, title, text) {
  return api.showDialog('info', title, { text }, [{ label: 'OK', default: true }], 'll-integration-message');
}

function openPath(api, targetPath, label) {
  if (!targetPath) {
    return showMessage(api, 'LL Integration', `${label} is not configured yet.`);
  }
  return shell.openPath(targetPath).then((error) => {
    if (error) {
      return showMessage(api, 'LL Integration', `${label} could not be opened:\n${error}`);
    }
    return undefined;
  });
}

function launchVortexManager(api, mode = 'links') {
  const appPath = nativeAppPath();
  const exePath = path.join(appPath, 'll_integration_vortex_manager.exe');
  let child;
  try {
    if (fs.existsSync(exePath)) {
      child = require('child_process').spawn(exePath, [`--mode=${mode}`], {
        detached: true,
        stdio: 'ignore',
      });
    } else {
      return showMessage(
        api,
        'LL Integration',
        `Vortex manager executable was not found in:\n${appPath}\n\nRun the LL Integration installer again or build the release package so PyQt6 is bundled inside the manager exe.`,
      );
    }
    child.unref();
    return undefined;
  } catch (err) {
    return showMessage(api, 'LL Integration', `Could not open Vortex manager:\n${err.message || err}`);
  }
}

function showTools(api) {
  const configPath = nativeConfigPath();
  const config = readJson(configPath);
  const metadataPath = config.metadata_path || path.join(nativeAppPath(), 'metadata');
  const vortexDownloads = config.vortex_downloads_path || '';
  const mo2Downloads = config.mo2_downloads_path || '';
  const activeTarget = config.active_downloads_target || 'mo2';

  const text = [
    `Active browser capture target: ${activeTarget}`,
    `Vortex downloads: ${vortexDownloads || '(not configured)'}`,
    `MO2 downloads: ${mo2Downloads || '(not configured)'}`,
  ].join('\n');

  return api.showDialog(
    'question',
    'LL Integration',
    { text },
    [
      { label: 'Open Vortex downloads', default: true },
      { label: 'Open metadata' },
      { label: 'Open native app' },
      { label: 'Open config' },
      { label: 'Close' },
    ],
    'll-integration-vortex-tools',
  ).then((result) => {
    switch (result.action) {
      case 'Open Vortex downloads':
        return openPath(api, vortexDownloads, 'Vortex downloads');
      case 'Open metadata':
        return openPath(api, metadataPath, 'Metadata folder');
      case 'Open native app':
        return openPath(api, nativeAppPath(), 'Native app folder');
      case 'Open config':
        return openPath(api, configPath, 'Native config');
      default:
        return undefined;
    }
  });
}

function showToolComingSoon(api, toolName) {
  return showMessage(
    api,
    'LL Integration',
    `${toolName} is currently available from MO2 Tools > LL Integration.\n\nThe Vortex page is wired for this tool, but the standalone Vortex implementation is not enabled yet.`,
  );
}

function LLIntegrationPage(props) {
  const rows = [
    ['Active browser capture target', props.activeTarget],
    ['Vortex downloads', props.vortexDownloads || '(not configured)'],
    ['MO2 downloads', props.mo2Downloads || '(not configured)'],
    ['Config', props.configPath],
  ];
  const toolButtonStyle = {
    display: 'block',
    width: '100%',
    margin: '8px 0',
    padding: '8px 12px',
    textAlign: 'center',
  };

  return React.createElement(
    vortexApi.MainPage,
    null,
    React.createElement(
      vortexApi.MainPage.Header,
      null,
      React.createElement('h1', null, 'LL Integration'),
    ),
    React.createElement(
      vortexApi.MainPage.Body,
      null,
      React.createElement(
        'div',
        { style: { padding: '24px', maxWidth: '920px' } },
        React.createElement('p', null, 'Manage LL Integration tools and check the current browser capture target.'),
        React.createElement(
          'div',
          { style: { maxWidth: '520px', margin: '18px 0 24px 0' } },
          React.createElement('button', { className: 'btn btn-primary', style: toolButtonStyle, onClick: props.manageLinks }, 'Manage LoversLab Links'),
          React.createElement('button', { className: 'btn btn-primary', style: toolButtonStyle, onClick: props.findVoicePacks }, 'Find Voice Packs'),
          React.createElement('button', { className: 'btn btn-primary', style: toolButtonStyle, onClick: props.createManualLink }, 'Create Manual Link'),
          React.createElement('button', { className: 'btn btn-primary', style: toolButtonStyle, onClick: props.purgeSuspiciousLinks }, 'Purge Suspicious Links'),
          React.createElement('button', { className: 'btn btn-primary', style: toolButtonStyle, onClick: props.integrationPaths }, 'Integration Paths'),
        ),
        React.createElement(
          'div',
          { style: { margin: '18px 0' } },
          rows.map(([label, value]) => React.createElement(
            'div',
            { key: label, style: { display: 'grid', gridTemplateColumns: '220px 1fr', gap: '12px', padding: '6px 0' } },
            React.createElement('strong', null, label),
            React.createElement('span', null, value),
          )),
        ),
        React.createElement(
          'div',
          { style: { display: 'flex', gap: '10px', flexWrap: 'wrap' } },
          React.createElement('button', { className: 'btn btn-default', onClick: props.syncMetadata }, 'Sync LL Metadata')
        ),
      ),
    ),
  );
}

function init(context) {
  const registrar = context.optional || context;
  const action = () => {
    showTools(context.api);
  };
  const isSupportedGame = () => {
    try {
      const state = context.api.store.getState();
      return SUPPORTED_GAME_IDS.has(vortexApi.selectors.activeGameId(state));
    } catch (err) {
      logLine('visible check failed', formatError(err));
      return false;
    }
  };

  registrar.registerMainPage('link', 'LL Integration', LLIntegrationPage, {
    id: 'll-integration',
    group: 'per-game',
    priority: 95,
    visible: isSupportedGame,
    props: () => {
      const configPath = nativeConfigPath();
      const config = readJson(configPath);
      return {
        activeTarget: config.active_downloads_target || 'mo2',
        vortexDownloads: config.vortex_downloads_path || '',
        mo2Downloads: config.mo2_downloads_path || '',
        configPath,
        manageLinks: () => launchVortexManager(context.api, 'links'),
        findVoicePacks: () => launchVortexManager(context.api, 'voice'),
        createManualLink: () => launchVortexManager(context.api, 'create-link'),
        purgeSuspiciousLinks: () => launchVortexManager(context.api, 'purge'),
        integrationPaths: () => showTools(context.api),
        syncMetadata: () => {
          const count = syncLLMetadata(context.api);
          writeVortexSnapshot(context.api);
          return showMessage(context.api, 'LL Integration', `Synced LL metadata for ${count} item(s).`);
        },
      };
    },
  });

  registrar.registerAction('global-icons', 100, 'show', {}, 'LL Integration', action);
  registrar.registerAction('mod-icons', 900, 'link', {}, 'LL Integration', action);
  registrar.registerAction('mods-action-icons', 900, 'link', {}, 'LL Integration', action);
  registrar.registerAction('download-actions', 900, 'link', {}, 'LL Integration', action);
  registrar.registerAction('download-icons', 900, 'link', {}, 'LL Integration', action);

  context.once(() => {
    try {
      logLine('extension once start');

      writeVortexSnapshot(context.api);

      // IMPORTANT:
      // Keep snapshot refresh, but do NOT call syncLLMetadata here.
      // syncLLMetadata dispatches Redux actions and can create a dispatch loop.
      context.api.store.subscribe(() => {
        clearTimeout(snapshotTimer);
        snapshotTimer = setTimeout(() => {
          try {
            writeVortexSnapshot(context.api);
          } catch (err) {
            logLine('snapshot refresh failed', formatError(err));
          }
        }, 1000);
      });

      context.api.events.on('did-install-mod', (eventGameId, archiveId, modId) => {
        try {
          const key = String(archiveId || '');
          const pending = pendingLLInstalls.get(key);

          logLine('did-install-mod received', {
            eventGameId,
            archiveId,
            modId,
            hasPendingLLInstall: !!pending,
          });

          if (pending) {
            const config = readJson(nativeConfigPath());
            const sidecar = findLLSidecarForArchive(config, pending.archivePath);
            const metadata = sidecar ? parseLLIni(sidecar) : {};

            const hasSidecarUrl = !!(
              metadata.page_url
              || metadata.mod_homepage
              || metadata.homepage
              || metadata.website
              || metadata.url
            );

            logLine('did-install-mod pending LL metadata source', {
              archiveId,
              modId,
              archivePath: pending.archivePath,
              sidecar,
              hasSidecarUrl,
              commandSourceUrl: pending.command.sourceUrl || pending.command.pageUrl || pending.command.modHomepage || '',
            });

            if (hasSidecarUrl) {
              applyLLMetadataToDownload(context.api, archiveId, pending.gameId || eventGameId, pending.archivePath, metadata);
              applyLLMetadataToInstalledMod(context.api, pending.gameId || eventGameId, modId, pending.archivePath, metadata);
            } else {
              applyInstalledModMetadata(
                context.api,
                pending.gameId || eventGameId,
                modId,
                pending.archivePath,
                pending.command,
              );
            }

            pendingLLInstalls.delete(key);
            writeVortexSnapshot(context.api);
            return;
          }

          const config = readJson(nativeConfigPath());

          // archiveId is usually Vortex's download id, not the real archive filename.
          // Resolve it through persistent.downloads.files first.
          const archivePath =
            downloadArchivePathFromId(context.api, archiveId)
            || archiveId;

          const sidecar = findLLSidecarForArchive(config, archivePath);
          if (!sidecar) {
            logLine('did-install-mod no LL sidecar found', {
              archiveId,
              resolvedArchivePath: archivePath,
            });
            return;
          }

          const metadata = parseLLIni(sidecar);
          if (!metadata.page_url && !metadata.download_url) {
            logLine('did-install-mod LL sidecar has no URL metadata', {
              archiveId,
              resolvedArchivePath: archivePath,
              sidecar,
            });
            return;
          }

          applyLLMetadataToDownload(context.api, archiveId, eventGameId, archivePath, metadata);
          applyLLMetadataToInstalledMod(context.api, eventGameId, modId, archivePath, metadata);
          writeVortexSnapshot(context.api);

        } catch (err) {
          logLine('did-install-mod LL metadata apply failed', formatError(err));
        }
      });

      startCommandPoller(context.api);

      logLine('extension once ready');
    } catch (err) {
      logLine('extension once failed', formatError(err));
    }
  });

  return true;
}

exports.default = init;
