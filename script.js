document.documentElement.classList.add('js-enabled');

const observer = new IntersectionObserver((entries) => {
  entries.forEach((entry) => {
    if (entry.isIntersecting) {
      entry.target.classList.add('visible');
      observer.unobserve(entry.target);
    }
  });
}, { threshold: 0.12 });

document.querySelectorAll('.reveal').forEach((element) => observer.observe(element));

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
