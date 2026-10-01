document.documentElement.classList.add('js-enabled');

const WORKBOOK_URL = 'portfolio-data.xlsx';
const LIST_SEPARATOR = '|';
let revealObserver = null;

function splitList(value) {
  return String(value || '').split(LIST_SEPARATOR).map((item) => item.trim()).filter(Boolean);
}

function truthy(value) {
  return ['yes', 'true', '1'].includes(String(value || '').trim().toLowerCase());
}

function makeElement(tag, className = '', text = null) {
  const element = document.createElement(tag);
  if (className) element.className = className;
  if (text !== null && text !== undefined) element.textContent = String(text);
  return element;
}

function observeReveals() {
  if (revealObserver) revealObserver.disconnect();
  revealObserver = new IntersectionObserver((entries) => {
    entries.forEach((entry) => {
      if (entry.isIntersecting) {
        entry.target.classList.add('visible');
        revealObserver.unobserve(entry.target);
      }
    });
  }, { threshold: 0.12 });
  document.querySelectorAll('.reveal').forEach((element) => revealObserver.observe(element));
}

function setupFilter(buttonSelector, itemSelector, categoryAttribute, emptyStateId) {
  const buttons = [...document.querySelectorAll(buttonSelector)];
  const items = [...document.querySelectorAll(itemSelector)];
  const emptyState = emptyStateId ? document.getElementById(emptyStateId) : null;
  if (!buttons.length || !items.length) return;

  buttons.forEach((button) => {
    button.addEventListener('click', () => {
      const filter = button.dataset.projectFilter || button.dataset.skillFilter || 'all';
      let visibleCount = 0;
      buttons.forEach((candidate) => {
        const active = candidate === button;
        candidate.classList.toggle('is-active', active);
        candidate.setAttribute('aria-pressed', String(active));
      });
      items.forEach((item) => {
        const categories = (item.getAttribute(categoryAttribute) || '').split(/\s+/).filter(Boolean);
        const visible = filter === 'all' || categories.includes(filter);
        item.classList.toggle('is-filtered-out', !visible);
        if (visible) visibleCount += 1;
      });
      if (emptyState) emptyState.hidden = visibleCount !== 0;
    });
  });
}

function setupInteractions() {
  setupFilter('[data-project-filter]', '[data-project-categories]', 'data-project-categories', 'project-empty');
  setupFilter('[data-skill-filter]', '[data-skill-category]', 'data-skill-category');
  document.querySelectorAll('.project-toggle').forEach((button) => {
    button.addEventListener('click', () => {
      const project = button.closest('.project');
      if (!project) return;
      const expanded = !project.classList.contains('is-expanded');
      project.classList.toggle('is-expanded', expanded);
      button.setAttribute('aria-expanded', String(expanded));
      button.textContent = expanded ? 'Less detail' : 'More detail';
    });
  });
  observeReveals();
}

function rowsFromSheet(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) throw new Error(`Missing required sheet: ${sheetName}`);
  return XLSX.utils.sheet_to_json(sheet, { defval: '' });
}

function renderProfile(rows) {
  const values = Object.fromEntries(rows.map((row) => [String(row.key), String(row.value || '')]));
  const eyebrow = document.querySelector('.hero .eyebrow');
  if (eyebrow && values.eyebrow) eyebrow.replaceChildren(makeElement('span', 'pulse'), document.createTextNode(` ${values.eyebrow}`));
  const heading = document.querySelector('.hero-copy h1');
  if (heading && values.hero_highlight) heading.replaceChildren(document.createTextNode(values.hero_prefix || ''), makeElement('span', '', values.hero_highlight), document.createTextNode(values.hero_suffix || ''));
  const lead = document.querySelector('.hero-lead');
  if (lead && values.hero_lead) lead.textContent = values.hero_lead;
  const signals = document.querySelector('.signal-row');
  if (signals) signals.replaceChildren(...splitList(values.hero_signals).map((item) => makeElement('span', '', item)));
  const chipName = document.querySelector('.profile-chip strong');
  const chipLine = document.querySelector('.profile-chip span');
  if (chipName && values.name) chipName.textContent = values.name;
  if (chipLine && values.profile_line) chipLine.textContent = values.profile_line;
  const footerName = document.querySelector('.site-footer strong');
  const footerRoles = document.querySelector('.site-footer div span');
  if (footerName && values.name) footerName.textContent = values.name;
  if (footerRoles && values.roles) footerRoles.textContent = splitList(values.roles).join(' · ');
  const ctaTitle = document.querySelector('.cta h2');
  const ctaBody = document.querySelector('.cta p:not(.kicker)');
  if (ctaTitle && values.cta_title) ctaTitle.textContent = values.cta_title;
  if (ctaBody && values.cta_body) ctaBody.textContent = values.cta_body;
}

function renderArchitecture(rows) {
  const map = document.querySelector('.architecture-map');
  const proof = document.querySelector('.console-proof');
  if (!map || !proof) return;
  const nodes = rows.filter((row) => row.kind === 'node').sort((a, b) => Number(a.order) - Number(b.order));
  const elements = [];
  nodes.forEach((row, index) => {
    const node = makeElement('div', `architecture-node node-${row.class_name}`);
    node.append(makeElement('span', '', row.code), makeElement('strong', '', row.title), makeElement('small', '', row.detail));
    elements.push(node);
    if (index < nodes.length - 1) {
      const link = makeElement('div', 'architecture-link');
      link.setAttribute('aria-hidden', 'true');
      link.append(makeElement('i'));
      elements.push(link);
    }
  });
  map.replaceChildren(...elements);
  proof.replaceChildren(...rows.filter((row) => row.kind === 'proof').sort((a,b)=>Number(a.order)-Number(b.order)).map((row) => makeElement('span', '', row.title)));
}

function renderImpact(rows) {
  const container = document.querySelector('.impact-strip');
  if (!container) return;
  container.replaceChildren(...rows.sort((a,b)=>Number(a.order)-Number(b.order)).map((row) => {
    const article = makeElement('article');
    article.append(makeElement('strong', '', row.title), makeElement('span', '', row.text));
    return article;
  }));
}

function renderCapabilities(rows) {
  const container = document.querySelector('.capability-atlas');
  if (!container) return;
  container.replaceChildren(...rows.sort((a,b)=>Number(a.order)-Number(b.order)).map((row, index) => {
    const article = makeElement('article', `capability-card reveal${index % 3 === 1 ? ' delay-1' : index % 3 === 2 ? ' delay-2' : ''}`);
    const head = makeElement('div', 'capability-head');
    head.append(makeElement('span', '', row.eyebrow), makeElement('strong', '', row.title));
    const stack = makeElement('div', 'capability-stack');
    splitList(row.tools).forEach((tool) => stack.append(makeElement('span', '', tool)));
    article.append(makeElement('div', 'capability-icon', row.icon), head, makeElement('p', '', row.description), stack);
    return article;
  }));
}

function renderDomains(rows) {
  const container = document.querySelector('.domain-strip');
  if (container) container.replaceChildren(...rows.sort((a,b)=>Number(a.order)-Number(b.order)).map((row) => makeElement('span', '', row.name)));
}

function renderProjectFilters(rows) {
  const container = document.querySelector('#work .filter-bar');
  if (!container) return;
  container.replaceChildren(...rows.sort((a,b)=>Number(a.order)-Number(b.order)).map((row, index) => {
    const button = makeElement('button', `filter-button${index === 0 ? ' is-active' : ''}`, row.label);
    button.type = 'button';
    button.dataset.projectFilter = row.key;
    button.setAttribute('aria-pressed', String(index === 0));
    return button;
  }));
}

function renderProjects(rows) {
  const container = document.getElementById('project-grid');
  if (!container) return;
  container.replaceChildren(...rows.sort((a,b)=>Number(a.order)-Number(b.order)).map((row) => {
    const isPublic = row.kind === 'public';
    const classes = ['project'];
    if (truthy(row.flagship)) classes.push('flagship');
    if (truthy(row.featured)) classes.push('featured');
    if (truthy(row.private)) classes.push('private');
    classes.push('reveal');
    if (String(row.delay) === '1') classes.push('delay-1');
    const item = makeElement(isPublic ? 'a' : 'article', classes.join(' '));
    item.setAttribute('data-project-categories', splitList(row.categories).join(' '));
    if (isPublic && row.url) { item.href = row.url; item.target = '_blank'; item.rel = 'noreferrer'; }
    const intro = makeElement('div');
    intro.append(makeElement('span', 'project-label', row.label), makeElement('h3', '', row.title), makeElement('p', '', row.summary));
    item.append(intro);
    const flow = splitList(row.flow);
    if (flow.length) {
      const panel = makeElement('div', 'evidence-panel flagship-flow');
      panel.setAttribute('aria-label', 'Sanitized architecture snapshot');
      panel.append(makeElement('em', '', 'Architecture flow'));
      flow.forEach((step, index) => {
        panel.append(makeElement('span', '', step));
        if (index < flow.length - 1) panel.append(makeElement('b', '', '→'));
      });
      item.append(panel);
    }
    if (!isPublic) {
      const details = makeElement('div', 'project-details');
      [['Problem:', row.problem], ['Architecture:', row.architecture], ['Value:', row.value]].forEach(([label, value], index) => {
        if (!value) return;
        details.append(makeElement('strong', '', label), document.createTextNode(` ${value}`));
        if (index < 2) details.append(document.createElement('br'));
      });
      item.append(details);
      if (row.evidence) {
        const note = makeElement('div', 'evidence-note');
        note.append(makeElement('strong', '', 'Evidence:'), document.createTextNode(` ${row.evidence}`));
        item.append(note);
      }
      const toggle = makeElement('button', 'project-toggle', 'More detail');
      toggle.type = 'button';
      toggle.setAttribute('aria-expanded', 'false');
      item.append(toggle);
    }
    const meta = makeElement('div', 'project-meta');
    splitList(row.tags).forEach((tag) => meta.append(makeElement('span', '', tag)));
    if (isPublic && row.url) meta.append(makeElement('strong', '', 'Open repo ↗'));
    item.append(meta);
    return item;
  }));
}

function renderExperience(rows) {
  const container = document.querySelector('.timeline');
  if (!container) return;
  container.replaceChildren(...rows.sort((a,b)=>Number(a.order)-Number(b.order)).map((row) => {
    const article = makeElement('article', 'timeline-item reveal');
    const content = makeElement('div');
    content.append(makeElement('h3', '', row.role), makeElement('p', 'timeline-company', row.company), makeElement('p', '', row.description));
    article.append(makeElement('div', 'timeline-date', row.dates), content);
    return article;
  }));
}

function renderLifecycle(rows) {
  const container = document.querySelector('.process');
  if (!container) return;
  container.replaceChildren(...rows.sort((a,b)=>Number(a.order)-Number(b.order)).map((row) => {
    const step = makeElement('div', 'step');
    const content = makeElement('div');
    content.append(makeElement('h3', '', row.title), makeElement('p', '', row.description));
    step.append(makeElement('span', '', String(row.order).padStart(2, '0')), content);
    return step;
  }));
}

function renderTechnologyFilters(rows) {
  const container = document.querySelector('#skills .filter-bar');
  if (!container) return;
  container.replaceChildren(...rows.sort((a,b)=>Number(a.order)-Number(b.order)).map((row, index) => {
    const button = makeElement('button', `filter-button${index === 0 ? ' is-active' : ''}`, row.label);
    button.type = 'button';
    button.dataset.skillFilter = row.key;
    button.setAttribute('aria-pressed', String(index === 0));
    return button;
  }));
}

function renderTechnology(rows) {
  const container = document.getElementById('skill-groups');
  if (!container) return;
  container.replaceChildren(...rows.sort((a,b)=>Number(a.order)-Number(b.order)).map((row) => {
    const article = makeElement('article', 'tool-group reveal');
    article.dataset.skillCategory = row.category;
    article.append(makeElement('h3', '', row.title));
    const stack = makeElement('div', 'stack');
    splitList(row.tools).forEach((tool) => stack.append(makeElement('span', '', tool)));
    article.append(stack);
    return article;
  }));
}

function renderCredentials(rows, educationRows) {
  const container = document.querySelector('.credentials-grid');
  if (container) {
    container.replaceChildren(...rows.sort((a,b)=>Number(a.order)-Number(b.order)).map((row) => {
      const article = makeElement('article');
      article.append(makeElement('strong', '', row.issuer), makeElement('span', '', row.credential));
      return article;
    }));
  }
  const education = document.querySelector('.education-note');
  const row = educationRows[0];
  if (education && row) education.textContent = `Education: ${row.qualification}, ${row.institution}, ${row.period}.`;
}

async function loadWorkbookContent() {
  if (!window.XLSX) throw new Error('SheetJS parser is unavailable');
  const response = await fetch(WORKBOOK_URL, { cache: 'no-cache' });
  if (!response.ok) throw new Error(`Workbook request failed with ${response.status}`);
  const workbook = XLSX.read(await response.arrayBuffer(), { type: 'array' });
  renderProfile(rowsFromSheet(workbook, 'Profile'));
  renderArchitecture(rowsFromSheet(workbook, 'Architecture'));
  renderImpact(rowsFromSheet(workbook, 'Impact'));
  renderCapabilities(rowsFromSheet(workbook, 'Capabilities'));
  renderDomains(rowsFromSheet(workbook, 'Domains'));
  renderProjectFilters(rowsFromSheet(workbook, 'ProjectFilters'));
  renderProjects(rowsFromSheet(workbook, 'Projects'));
  renderExperience(rowsFromSheet(workbook, 'Experience'));
  renderLifecycle(rowsFromSheet(workbook, 'Lifecycle'));
  renderTechnologyFilters(rowsFromSheet(workbook, 'TechnologyFilters'));
  renderTechnology(rowsFromSheet(workbook, 'Technology'));
  renderCredentials(rowsFromSheet(workbook, 'Credentials'), rowsFromSheet(workbook, 'Education'));
  document.documentElement.dataset.contentSource = 'excel';
}

loadWorkbookContent().then(setupInteractions).catch((error) => {
  document.documentElement.dataset.contentSource = 'fallback';
  console.warn('Portfolio workbook enhancement failed; static fallback content remains active.', error);
  setupInteractions();
});
