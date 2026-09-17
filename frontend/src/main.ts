import './style.css';
import {
  ArrowRight,
  Check,
  CheckCircle2,
  ChevronDown,
  Circle,
  ExternalLink,
  FileArchive,
  FolderOpen,
  HardDrive,
  Info,
  Laptop,
  Mail,
  Moon,
  Pause,
  Play,
  RefreshCw,
  RotateCcw,
  ShieldCheck,
  Sun,
  TriangleAlert,
  XCircle,
  createIcons,
} from 'lucide';
import {
  CancelMigration,
  GetEnvironmentStatus,
  OpenFolder,
  SelectStagingFolder,
  SelectThunderbirdFolder,
  StartMigration,
} from '../wailsjs/go/main/App';
import { EventsOn } from '../wailsjs/runtime/runtime';

interface EnvironmentStatus {
  windows: boolean;
  outlookRunning: boolean;
  pstNames: string[];
  message: string;
}

interface MigrationProgress {
  stage: string;
  status: 'idle' | 'running' | 'complete' | 'error' | 'cancelled';
  message: string;
  folder: string;
  current: number;
  total: number;
  converted: number;
  imported: number;
  failed: number;
}

const state = {
  environment: null as EnvironmentStatus | null,
  progress: {
    stage: 'prepare',
    status: 'idle',
    message: 'Complete the setup to begin.',
    folder: '',
    current: 0,
    total: 0,
    converted: 0,
    imported: 0,
    failed: 0,
  } as MigrationProgress,
  sourcePath: '',
  stagingPath: '',
  targetPSTName: '',
  markAsRead: true,
  preventSleep: true,
  keepStagingEML: false,
  skipConversion: false,
};

const app = document.querySelector<HTMLDivElement>('#app');
if (!app) throw new Error('Application root was not found');

app.innerHTML = `
  <div class="min-h-screen">
    <header class="border-b border-base-300/80 bg-base-100/90 backdrop-blur">
      <div class="mx-auto flex max-w-7xl items-center justify-between px-6 py-4">
        <div class="flex items-center gap-3">
          <div class="grid size-11 place-items-center rounded-box bg-primary text-primary-content shadow-sm">
            <i data-lucide="mail" class="size-6"></i>
          </div>
          <div>
            <h1 class="text-xl font-bold tracking-tight">Mailbox Mover</h1>
            <p class="text-sm text-base-content/60">Thunderbird to Classic Outlook</p>
          </div>
        </div>
        <div class="flex items-center gap-2">
          <span id="platform-badge" class="badge badge-soft">Checking Outlook</span>
          <button id="theme-button" class="btn btn-ghost btn-circle" aria-label="Toggle colour theme">
            <i data-lucide="moon" class="size-5"></i>
          </button>
        </div>
      </div>
    </header>

    <main class="mx-auto grid max-w-7xl gap-6 px-6 py-7 lg:grid-cols-[minmax(0,1fr)_21rem]">
      <section class="min-w-0 space-y-6" aria-labelledby="migration-title">
        <div>
          <p class="mb-2 text-sm font-semibold uppercase tracking-[0.16em] text-primary">Thunderbird to Outlook</p>
          <h2 id="migration-title" class="text-3xl font-bold tracking-tight">Move Thunderbird mail into Outlook</h2>
          <p class="mt-2 max-w-2xl text-base-content/65">Choose the mail source, an Outlook data file, and a folder for EML files. Keep Classic Outlook open during the import.</p>
        </div>

        <ul id="migration-steps" class="steps steps-horizontal w-full text-xs sm:text-sm">
          <li class="step step-primary">Source</li>
          <li class="step">Outlook</li>
          <li class="step">Review</li>
          <li class="step">Move</li>
        </ul>

        <div class="card card-border bg-base-100 shadow-sm">
          <div class="card-body gap-5">
            <div class="flex items-start gap-4">
              <div class="grid size-11 shrink-0 place-items-center rounded-box bg-primary/10 text-primary"><i data-lucide="folder-open" class="size-6"></i></div>
              <div class="min-w-0 grow">
                <div class="flex flex-wrap items-center justify-between gap-2">
                  <div>
                    <h3 class="card-title text-lg">1. Choose Thunderbird mail</h3>
                    <p id="source-description" class="mt-1 text-sm text-base-content/60">Select the <strong>Local Folders</strong> directory from your Thunderbird profile.</p>
                  </div>
                  <span id="source-badge" class="badge badge-ghost">Required</span>
                </div>
                <div class="mt-4 flex gap-2">
                  <label class="input min-w-0 grow">
                    <i data-lucide="folder-open" class="size-4 text-base-content/45"></i>
                    <input id="source-path" class="path-text" type="text" placeholder="No source folder selected" readonly />
                  </label>
                  <button id="source-button" class="btn"><i data-lucide="folder-open" class="size-4"></i>Browse</button>
                </div>
                <button id="source-help" class="mt-3 inline-flex items-center gap-1 text-sm text-primary hover:underline"><i data-lucide="info" class="size-4"></i>Where is this folder?</button>
              </div>
            </div>
          </div>
        </div>

        <div class="card card-border bg-base-100 shadow-sm">
          <div class="card-body gap-5">
            <div class="flex items-start gap-4">
              <div class="grid size-11 shrink-0 place-items-center rounded-box bg-secondary/20 text-secondary-content"><i data-lucide="hard-drive" class="size-6"></i></div>
              <div class="min-w-0 grow">
                <div class="flex flex-wrap items-center justify-between gap-2">
                  <div>
                    <h3 class="card-title text-lg">2. Choose your Outlook archive</h3>
                    <p class="mt-1 text-sm text-base-content/60">Classic Outlook must be open with the destination PST already attached.</p>
                  </div>
                  <button id="refresh-outlook" class="btn btn-ghost btn-sm"><i data-lucide="refresh-cw" class="size-4"></i>Check again</button>
                </div>
                <fieldset class="fieldset mt-3">
                  <legend class="fieldset-legend">Outlook data file</legend>
                  <select id="pst-select" class="select w-full" disabled>
                    <option value="">Open Classic Outlook to load available PSTs</option>
                  </select>
                  <p id="outlook-message" class="label">Checking Classic Outlook.</p>
                </fieldset>
              </div>
            </div>
          </div>
        </div>

        <div class="card card-border bg-base-100 shadow-sm">
          <div class="card-body gap-5">
            <div class="flex items-start gap-4">
              <div class="grid size-11 shrink-0 place-items-center rounded-box bg-accent/20 text-accent-content"><i data-lucide="file-archive" class="size-6"></i></div>
              <div class="min-w-0 grow">
                <h3 id="folder-title" class="card-title text-lg">3. Temporary working folder</h3>
                <p id="folder-description" class="mt-1 text-sm text-base-content/60">The app writes converted EML files here before Outlook imports them.</p>
                <div class="mt-4 flex gap-2">
                  <label class="input min-w-0 grow">
                    <i data-lucide="file-archive" class="size-4 text-base-content/45"></i>
                    <input id="staging-path" class="path-text" type="text" placeholder="Choose an empty folder with enough free space" readonly />
                  </label>
                  <button id="staging-button" class="btn"><i data-lucide="folder-open" class="size-4"></i>Browse</button>
                </div>
              </div>
            </div>
          </div>
        </div>

        <div class="card card-border bg-base-100 shadow-sm">
          <div class="card-body">
            <h3 class="card-title text-lg">Import preferences</h3>
            <div class="mt-2 divide-y divide-base-300">
              <label class="flex cursor-pointer items-center justify-between gap-4 py-4">
                <span><span class="font-medium">Mark imported messages as read</span><span class="mt-1 block text-sm text-base-content/55">Recommended for historical mail.</span></span>
                <input id="mark-read" type="checkbox" class="toggle" checked />
              </label>
              <label class="flex cursor-pointer items-center justify-between gap-4 py-4">
                <span><span class="font-medium">Keep this computer awake</span><span class="mt-1 block text-sm text-base-content/55">Prevents sleep while a long import is running.</span></span>
                <input id="prevent-sleep" type="checkbox" class="toggle" checked />
              </label>
              <label class="flex cursor-pointer items-center justify-between gap-4 py-4">
                <span><span class="font-medium">Keep temporary EML files</span><span class="mt-1 block text-sm text-base-content/55">Useful as a backup, but uses more disk space.</span></span>
                <input id="keep-eml" type="checkbox" class="toggle" />
              </label>
              <label class="flex cursor-pointer items-center justify-between gap-4 py-4">
                <span><span class="font-medium">My source already contains EML files</span><span class="mt-1 block text-sm text-base-content/55">Skip Thunderbird mbox conversion and import the working folder directly.</span></span>
                <input id="skip-conversion" type="checkbox" class="toggle" />
              </label>
            </div>
          </div>
        </div>

        <div id="progress-card" class="card card-border hidden bg-base-100 shadow-sm" aria-live="polite">
          <div class="card-body">
            <div class="flex items-start justify-between gap-4">
              <div>
                <div class="flex items-center gap-2"><span id="progress-spinner" class="loading loading-spinner loading-sm"></span><h3 id="progress-title" class="card-title text-lg">Preparing migration</h3></div>
                <p id="progress-message" class="mt-2 text-sm text-base-content/65">Checking the selected folders.</p>
              </div>
              <span id="progress-percent" class="badge badge-soft">0%</span>
            </div>
            <progress id="progress-bar" class="progress progress-primary mt-3 w-full" value="0" max="100"></progress>
            <div class="stats stats-horizontal mt-3 w-full border border-base-300 bg-base-200/60">
              <div class="stat py-3"><div class="stat-title">Converted</div><div id="converted-count" class="stat-value text-2xl">0</div></div>
              <div class="stat py-3"><div class="stat-title">Imported</div><div id="imported-count" class="stat-value text-2xl">0</div></div>
              <div class="stat py-3"><div class="stat-title">Skipped</div><div id="failed-count" class="stat-value text-2xl">0</div></div>
            </div>
            <div class="card-actions mt-2 justify-end">
              <button id="open-output" class="btn btn-ghost btn-sm hidden"><i data-lucide="external-link" class="size-4"></i>Open EML folder</button>
              <button id="cancel-button" class="btn btn-outline btn-sm"><i data-lucide="pause" class="size-4"></i>Stop after current message</button>
            </div>
          </div>
        </div>

        <div id="form-alert" role="alert" class="alert alert-error hidden"></div>

        <div class="flex flex-col-reverse items-stretch justify-between gap-3 pb-4 sm:flex-row sm:items-center">
          <p class="flex items-center gap-2 text-sm text-base-content/55"><i data-lucide="shield-check" class="size-4"></i>Your original Thunderbird files are never modified.</p>
          <button id="start-button" class="btn btn-primary btn-lg min-w-56"><i data-lucide="play" class="size-5"></i>Start migration<i data-lucide="arrow-right" class="size-5"></i></button>
        </div>
      </section>

      <aside class="space-y-4 lg:sticky lg:top-6 lg:self-start" aria-label="Readiness summary">
        <div class="card card-border bg-base-100 shadow-sm">
          <div class="card-body">
            <h3 class="card-title text-lg">Setup status</h3>
            <ul class="list mt-2">
              <li class="list-row px-0"><span id="check-source-icon" class="text-base-content/30"><i data-lucide="circle" class="size-5"></i></span><div><div class="font-medium">Thunderbird source</div><div id="check-source" class="text-xs text-base-content/55">Not selected</div></div></li>
              <li class="list-row px-0"><span id="check-outlook-icon" class="text-base-content/30"><i data-lucide="circle" class="size-5"></i></span><div><div class="font-medium">Classic Outlook</div><div id="check-outlook" class="text-xs text-base-content/55">Checking</div></div></li>
              <li class="list-row px-0"><span id="check-pst-icon" class="text-base-content/30"><i data-lucide="circle" class="size-5"></i></span><div><div class="font-medium">Destination PST</div><div id="check-pst" class="text-xs text-base-content/55">Not selected</div></div></li>
              <li class="list-row px-0"><span id="check-staging-icon" class="text-base-content/30"><i data-lucide="circle" class="size-5"></i></span><div><div class="font-medium">Working folder</div><div id="check-staging" class="text-xs text-base-content/55">Not selected</div></div></li>
            </ul>
          </div>
        </div>

        <div class="alert alert-info alert-soft items-start">
          <i data-lucide="info" class="mt-0.5 size-5 shrink-0"></i>
          <div>
            <p class="font-semibold">Before you begin</p>
            <ul class="mt-2 list-disc space-y-1 pl-4 text-sm">
              <li>Use Classic Outlook, not New Outlook.</li>
              <li>Create or attach the target PST first.</li>
              <li>Close Thunderbird for a consistent copy.</li>
            </ul>
          </div>
        </div>

        <details class="collapse-arrow collapse border border-base-300 bg-base-100 shadow-sm">
          <summary class="collapse-title font-semibold">Migration sequence</summary>
          <div class="collapse-content text-sm text-base-content/65">
            <ol class="list-decimal space-y-2 pl-4">
              <li>Read Thunderbird mbox files.</li>
              <li>Create temporary EML messages.</li>
              <li>Rebuild the folder tree in your PST.</li>
              <li>Import mail, dates, senders, and attachments.</li>
            </ol>
          </div>
        </details>
      </aside>
    </main>
  </div>
`;

createIcons({
  icons: {
    ArrowRight,
    Check,
    CheckCircle2,
    ChevronDown,
    Circle,
    ExternalLink,
    FileArchive,
    FolderOpen,
    HardDrive,
    Info,
    Laptop,
    Mail,
    Moon,
    Pause,
    Play,
    RefreshCw,
    RotateCcw,
    ShieldCheck,
    Sun,
    TriangleAlert,
    XCircle,
  },
});

const element = <T extends HTMLElement>(id: string) => {
  const found = document.getElementById(id) as T | null;
  if (!found) throw new Error(`Missing element: ${id}`);
  return found;
};

function icon(name: string, className = 'size-5'): string {
  return `<i data-lucide="${name}" class="${className}"></i>`;
}

function refreshIcons(): void {
  createIcons({
    icons: { ArrowRight, Check, CheckCircle2, ChevronDown, Circle, ExternalLink, FileArchive, FolderOpen, HardDrive, Info, Laptop, Mail, Moon, Pause, Play, RefreshCw, RotateCcw, ShieldCheck, Sun, TriangleAlert, XCircle },
  });
}

function basename(path: string): string {
  return path.split(/[\\/]/).filter(Boolean).pop() ?? path;
}

function setCheck(id: string, ready: boolean, text: string): void {
  element(`${id}-icon`).innerHTML = ready ? icon('check-circle-2') : icon('circle');
  element(`${id}-icon`).className = ready ? 'text-success' : 'text-base-content/30';
  element(id).textContent = text;
}

function render(): void {
  element<HTMLInputElement>('source-path').value = state.sourcePath;
  element<HTMLInputElement>('staging-path').value = state.stagingPath;
  element<HTMLInputElement>('mark-read').checked = state.markAsRead;
  element<HTMLInputElement>('prevent-sleep').checked = state.preventSleep;
  element<HTMLInputElement>('keep-eml').checked = state.keepStagingEML;
  element<HTMLInputElement>('skip-conversion').checked = state.skipConversion;

  element('source-badge').textContent = state.skipConversion ? 'Using EML folder' : state.sourcePath ? 'Selected' : 'Required';
  element('source-badge').className = state.sourcePath || state.skipConversion ? 'badge badge-success badge-soft' : 'badge badge-ghost';
  element('source-description').textContent = state.skipConversion
    ? 'Thunderbird conversion is off. Choose the folder that contains your EML files below.'
    : 'Select the Local Folders directory from your Thunderbird profile.';
  element('folder-title').textContent = state.skipConversion ? '3. EML source folder' : '3. Temporary working folder';
  element('folder-description').textContent = state.skipConversion
    ? 'Choose the folder that contains the EML files to import.'
    : 'The app writes converted EML files here before Outlook imports them.';
  element<HTMLInputElement>('staging-path').placeholder = state.skipConversion
    ? 'Choose the folder that contains your EML files'
    : 'Choose an empty folder with enough free space';

  const status = state.environment;
  const platformBadge = element('platform-badge');
  platformBadge.textContent = status?.outlookRunning ? 'Outlook ready' : status?.windows ? 'Outlook not connected' : 'Windows required';
  platformBadge.className = status?.outlookRunning ? 'badge badge-success badge-soft' : 'badge badge-warning badge-soft';
  element('outlook-message').textContent = status?.message ?? 'Checking Classic Outlook.';

  const select = element<HTMLSelectElement>('pst-select');
  const names = status?.pstNames ?? [];
  select.innerHTML = names.length
    ? `<option value="">Select an Outlook data file</option>${names.map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`).join('')}`
    : '<option value="">Open Classic Outlook to load available PSTs</option>';
  select.disabled = names.length === 0 || state.progress.status === 'running';
  if (names.includes(state.targetPSTName)) select.value = state.targetPSTName;

  setCheck('check-source', state.skipConversion || Boolean(state.sourcePath), state.skipConversion ? 'Using existing EML files' : state.sourcePath ? basename(state.sourcePath) : 'Not selected');
  setCheck('check-outlook', Boolean(status?.outlookRunning), status?.outlookRunning ? 'Connected' : 'Not connected');
  setCheck('check-pst', Boolean(state.targetPSTName), state.targetPSTName || 'Not selected');
  setCheck('check-staging', Boolean(state.stagingPath), state.stagingPath ? basename(state.stagingPath) : 'Not selected');

  const isRunning = state.progress.status === 'running';
  const isReady = Boolean((state.sourcePath || state.skipConversion) && state.stagingPath && state.targetPSTName && status?.outlookRunning);
  const start = element<HTMLButtonElement>('start-button');
  start.disabled = !isReady || isRunning;
  start.innerHTML = isRunning
    ? `${icon('refresh-cw', 'size-5 animate-spin')}Moving messages`
    : state.progress.status === 'complete'
      ? `${icon('rotate-ccw', 'size-5')}Start another migration`
      : `${icon('play', 'size-5')}Start migration${icon('arrow-right', 'size-5')}`;

  for (const id of ['source-button', 'staging-button', 'refresh-outlook']) {
    element<HTMLButtonElement>(id).disabled = isRunning;
  }
  element<HTMLInputElement>('source-path').disabled = isRunning || state.skipConversion;

  renderProgress();
  renderSteps();
  refreshIcons();
}

function renderSteps(): void {
  let active = 0;
  if (state.sourcePath || state.skipConversion) active = 1;
  if (state.targetPSTName) active = 2;
  if (state.stagingPath && state.targetPSTName) active = 3;
  if (state.progress.status === 'running' || state.progress.status === 'complete') active = 4;
  element('migration-steps').querySelectorAll('.step').forEach((step, index) => {
    step.classList.toggle('step-primary', index < active || index === 0);
  });
}

function renderProgress(): void {
  const progress = state.progress;
  const visible = progress.status !== 'idle';
  element('progress-card').classList.toggle('hidden', !visible);
  if (!visible) return;

  const percent = progress.total > 0 ? Math.min(100, Math.round((progress.current / progress.total) * 100)) : progress.stage === 'convert' ? 20 : progress.stage === 'cleanup' ? 95 : progress.status === 'complete' ? 100 : 5;
  const titles: Record<string, string> = {
    prepare: 'Preparing migration',
    convert: 'Converting Thunderbird mail',
    import: 'Importing into Outlook',
    cleanup: 'Cleaning up',
    complete: 'Migration complete',
  };
  element('progress-title').textContent = progress.status === 'error' ? 'Migration could not continue' : progress.status === 'cancelled' ? 'Migration stopped' : titles[progress.stage] ?? 'Migration running';
  element('progress-message').textContent = progress.message;
  element('progress-percent').textContent = `${percent}%`;
  element<HTMLProgressElement>('progress-bar').value = percent;
  element('converted-count').textContent = progress.converted.toLocaleString();
  element('imported-count').textContent = progress.imported.toLocaleString();
  element('failed-count').textContent = progress.failed.toLocaleString();

  const spinner = element('progress-spinner');
  spinner.className = progress.status === 'running' ? 'loading loading-spinner loading-sm' : 'hidden';
  const cancel = element<HTMLButtonElement>('cancel-button');
  cancel.classList.toggle('hidden', progress.status !== 'running');
  element('open-output').classList.toggle('hidden', !(progress.status === 'complete' || progress.status === 'error' || progress.status === 'cancelled'));

  const card = element('progress-card');
  card.classList.toggle('border-error', progress.status === 'error');
  card.classList.toggle('border-success', progress.status === 'complete');
}

function escapeHtml(value: string): string {
  return value.replace(/[&<>'"]/g, (character) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' })[character] ?? character);
}

function showError(message: unknown): void {
  const alert = element('form-alert');
  alert.innerHTML = `${icon('triangle-alert')}<span>${escapeHtml(String(message))}</span>`;
  alert.classList.remove('hidden');
  refreshIcons();
}

function clearError(): void {
  const alert = element('form-alert');
  alert.className = 'alert alert-error hidden';
  alert.innerHTML = '';
}

async function refreshEnvironment(): Promise<void> {
  clearError();
  const button = element<HTMLButtonElement>('refresh-outlook');
  button.disabled = true;
  try {
    state.environment = await GetEnvironmentStatus() as EnvironmentStatus;
    const names = state.environment.pstNames ?? [];
    if (names.length === 1 && !state.targetPSTName) state.targetPSTName = names[0];
    if (state.targetPSTName && !names.includes(state.targetPSTName)) state.targetPSTName = '';
  } catch (error) {
    showError(error);
  } finally {
    button.disabled = false;
    render();
  }
}

element('source-button').addEventListener('click', async () => {
  clearError();
  try {
    const path = await SelectThunderbirdFolder();
    if (path) state.sourcePath = path;
    render();
  } catch (error) { showError(error); }
});

element('staging-button').addEventListener('click', async () => {
  clearError();
  try {
    const path = await SelectStagingFolder();
    if (path) state.stagingPath = path;
    render();
  } catch (error) { showError(error); }
});

element('source-help').addEventListener('click', () => {
  const alert = element('form-alert');
  alert.className = 'alert alert-info';
  alert.innerHTML = `${icon('info')}<span>Usually: <strong>%APPDATA%\\Thunderbird\\Profiles\\&lt;profile&gt;\\Mail\\Local Folders</strong>. In Thunderbird, open Account Settings, then select Local Folders to see the exact path.</span>`;
  refreshIcons();
});

element('refresh-outlook').addEventListener('click', refreshEnvironment);

element<HTMLSelectElement>('pst-select').addEventListener('change', (event) => {
  state.targetPSTName = (event.target as HTMLSelectElement).value;
  render();
});

for (const [id, key] of [
  ['mark-read', 'markAsRead'],
  ['prevent-sleep', 'preventSleep'],
  ['keep-eml', 'keepStagingEML'],
  ['skip-conversion', 'skipConversion'],
] as const) {
  element<HTMLInputElement>(id).addEventListener('change', (event) => {
    state[key] = (event.target as HTMLInputElement).checked;
    render();
  });
}

element('theme-button').addEventListener('click', () => {
  const root = document.documentElement;
  const dark = root.dataset.theme === 'mailbox-dark';
  root.dataset.theme = dark ? 'mailbox' : 'mailbox-dark';
  element('theme-button').innerHTML = icon(dark ? 'moon' : 'sun');
  refreshIcons();
});

element('start-button').addEventListener('click', async () => {
  clearError();
  state.progress = { stage: 'prepare', status: 'running', message: 'Checking the selected folders.', folder: '', current: 0, total: 0, converted: 0, imported: 0, failed: 0 };
  render();
  try {
    await StartMigration({
      sourcePath: state.sourcePath,
      stagingPath: state.stagingPath,
      targetPSTName: state.targetPSTName,
      markAsRead: state.markAsRead,
      preventSleep: state.preventSleep,
      skipConversion: state.skipConversion,
      keepStagingEML: state.keepStagingEML,
    });
  } catch (error) {
    state.progress.status = 'error';
    state.progress.message = String(error);
    showError(error);
    render();
  }
});

element('cancel-button').addEventListener('click', async () => {
  await CancelMigration();
  element('progress-message').textContent = 'The app will stop after the current message.';
});

element('open-output').addEventListener('click', async () => {
  try { await OpenFolder(state.stagingPath); } catch (error) { showError(error); }
});

EventsOn('migration:progress', (progress: MigrationProgress) => {
  state.progress = progress;
  render();
});

void refreshEnvironment();
render();
