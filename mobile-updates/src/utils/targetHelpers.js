/**
 * Updated targetHelpers.js for TrackSDN Mobile
 * Works with the new dynamic hierarchy API response format.
 *
 * New API response shape from GET /api/mobile/upload-targets:
 * {
 *   schema: { containerLabel, containerPlural, subContainerLabel, subContainerPlural, dependentLabel, dependentPlural },
 *   containers: [{ id, name, subContainers: [{ id, name, startDate, finishDate }] }],
 *   dependents: [{ id, name, supplier, permissions: { users: [] } }],
 *   rows: [{ pageId, pageName, rowId, packageName, supplierName, block, level, generated? }]
 * }
 */

export function unique(values) {
  const out = [];
  const seen = new Set();
  for (const value of values || []) {
    const x = String(value || '').trim();
    if (!x) continue;
    const key = x.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(x);
  }
  return out;
}

export function normaliseProjects(payload) {
  const state = payload?.state || payload || {};
  if (Array.isArray(payload?.projects)) return payload.projects;
  if (Array.isArray(state?.projects)) return state.projects;
  if (Array.isArray(payload)) return payload;
  return [];
}

export function getProjectId(project = {}) {
  return String(
    project.id || project.projectId || project.project_id ||
    project.projectNumber || project.number || project.name ||
    project.projectName || 'project'
  );
}

export function getProjectName(project = {}) {
  return String(project.name || project.projectName || project.title || 'Untitled project');
}

export function defaultSchema() {
  return {
    containerLabel: 'Block',
    containerPlural: 'Blocks',
    subContainerLabel: 'Level',
    subContainerPlural: 'Levels',
    dependentLabel: 'Package',
    dependentPlural: 'Packages',
  };
}

/**
 * Extracts upload targets from the new API response format.
 * Returns { schema, containers, dependents, rows, blocks, levels, packages, blockLevels }
 */
export function extractTargets(project = {}, apiTargets = null) {
  if (apiTargets) {
    const schema = { ...defaultSchema(), ...(apiTargets.schema || {}) };

    // New format: containers/subContainers/dependents
    const containers = Array.isArray(apiTargets.containers) ? apiTargets.containers : [];
    const dependents = Array.isArray(apiTargets.dependents) ? apiTargets.dependents : [];
    const rows = Array.isArray(apiTargets.rows) ? apiTargets.rows : [];

    // Flatten for legacy compatibility
    const blocks = unique(containers.map((c) => c.name));
    const levels = unique(containers.flatMap((c) => (c.subContainers || []).map((sc) => sc.name)));
    const blockLevels = containers.flatMap((c) =>
      (c.subContainers || []).map((sc) => ({ block: c.name, level: sc.name }))
    );
    const packages = unique(dependents.map((d) => d.name));

    return { schema, containers, dependents, rows, blocks, levels, blockLevels, packages };
  }

  // Fallback: old format or direct project object
  const schema = { ...defaultSchema(), ...(project.schema || {}) };
  const responsibilities = Array.isArray(project.responsibilities) ? project.responsibilities : [];
  const master = Array.isArray(project.master) ? project.master : [];
  const pages = Array.isArray(project.pages) ? project.pages : [];

  const containers = master.map((m) => ({
    id: m.id,
    name: String(m.blockZone || m.block || m.name || '').trim(),
    subContainers: (Array.isArray(m.levels) ? m.levels : []).map((lv) => ({
      id: lv.id,
      name: String(lv.name || lv.levelName || '').trim(),
      startDate: lv.startDate || '',
      finishDate: lv.finishDate || '',
    })),
  })).filter((c) => c.name);

  const dependents = responsibilities.map((r) => ({
    id: r.id,
    name: String(r.name || r.packageName || '').trim(),
    supplier: String(r.supplier || '').trim(),
    permissions: r.permissions || { users: [] },
  })).filter((d) => d.name);

  const blocks = unique(containers.map((c) => c.name));
  const levels = unique(containers.flatMap((c) => (c.subContainers || []).map((sc) => sc.name)));
  const blockLevels = containers.flatMap((c) =>
    (c.subContainers || []).map((sc) => ({ block: c.name, level: sc.name }))
  );
  const packages = unique(dependents.map((d) => d.name));

  const rows = pages
    .filter((pg) => !pg.meta?.isMaster)
    .flatMap((pg) => {
      const respId = pg.meta?.responsibilityId;
      const resp = responsibilities.find((r) => r.id === respId);
      const packageName = pg.meta?.packageName || resp?.name || pg.name || '';
      return (Array.isArray(pg.rows) ? pg.rows : [])
        .map((row) => ({
          pageId: pg.id,
          pageName: pg.name,
          rowId: row.id,
          packageName,
          supplierName: resp?.supplier || '',
          block: row.meta?.blockZone || row.blockZone || row.block || '',
          level: row.meta?.levelName || row.levelName || row.level || '',
        }))
        .filter((r) => r.block && r.level);
    });

  return { schema, containers, dependents, rows, blocks, levels, blockLevels, packages };
}
