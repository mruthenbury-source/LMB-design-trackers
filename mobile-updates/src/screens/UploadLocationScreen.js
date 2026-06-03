/**
 * Updated UploadLocationScreen.js for TrackSDN Mobile
 * Uses dynamic schema labels (containerLabel, subContainerLabel, dependentLabel)
 * returned from GET /api/mobile/upload-targets.
 */
import React, { useEffect, useMemo, useState } from 'react';
import { ActivityIndicator, Pressable, ScrollView, StyleSheet, Text, View } from 'react-native';
import Screen from '../components/Screen';
import AppButton from '../components/AppButton';
import Pill from '../components/Pill';
import { colors } from '../utils/theme';
import { defaultSchema, extractTargets, getProjectName } from '../utils/targetHelpers';
import { fetchUploadTargets } from '../services/api';

function clean(value) { return String(value || '').trim(); }
function key(value) { return clean(value).toLowerCase(); }

function looksLikeId(value = '') {
  const raw = clean(value);
  if (!raw) return false;
  if (/^[0-9a-f]{24,}$/i.test(raw)) return true;
  if (/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(raw)) return true;
  return false;
}

function firstNonId(...values) {
  for (const value of values) {
    const found = clean(value);
    if (found && !looksLikeId(found)) return found;
  }
  return '';
}

function numericAwareCompare(a = '', b = '') {
  const ax = clean(a); const bx = clean(b);
  const anum = ax.match(/-?\d+/); const bnum = bx.match(/-?\d+/);
  if (anum && bnum) { const diff = Number(anum[0]) - Number(bnum[0]); if (diff !== 0) return diff; }
  return ax.localeCompare(bx, undefined, { numeric: true, sensitivity: 'base' });
}

function compareByLevel(a = {}, b = {}) {
  const bc = numericAwareCompare(a.block, b.block);
  return bc !== 0 ? bc : numericAwareCompare(a.level, b.level);
}

function first(...values) { for (const v of values) { const f = clean(v); if (f) return f; } return ''; }

function get(row = {}, ...paths) {
  for (const path of paths) {
    const parts = String(path).split('.');
    let value = row;
    for (const part of parts) value = value?.[part];
    const found = clean(value);
    if (found) return found;
  }
  return '';
}

function rowIdOf(row = {}) { return first(get(row, 'id'), get(row, 'rowId'), get(row, 'key')); }
function pageIdOf(row = {}) { return first(get(row, 'pageId'), get(row, 'page.id'), get(row, 'pageName')); }
function pageNameOf(row = {}) { return first(get(row, 'pageName'), get(row, 'page.name'), 'Project Home'); }
function packageOf(row = {}) { return first(get(row, 'packageName'), get(row, 'package'), get(row, 'lineItemName'), get(row, 'item')); }
function blockOf(row = {}) { return first(get(row, 'block'), get(row, 'blockZone'), get(row, 'meta.blockZone')); }
function levelOf(row = {}) { return first(get(row, 'level'), get(row, 'levelName'), get(row, 'meta.levelName')); }
function supplierOf(row = {}) { return firstNonId(get(row, 'supplier'), get(row, 'supplierName'), get(row, 'meta.supplier')); }
function responsibilityIdOf(row = {}) { return first(get(row, 'responsibilityId'), get(row, 'packageId'), get(row, 'meta.responsibilityId')); }
function uploadScopeKeyOf(row = {}) { return first(get(row, 'uploadScopeKey'), get(row, 'meta.uploadScopeKey')); }

function pageMatchesName(page = {}, wanted = '') {
  const w = key(wanted);
  if (!w) return false;
  return [page?.name, page?.pageName, page?.packageName].some((v) => key(v) === w);
}

function pageIdForName(pages = [], wanted = '') {
  const m = pages.find((p) => pageMatchesName(p, wanted));
  return first(get(m || {}, 'id'), get(m || {}, 'pageId'), get(m || {}, 'name'), wanted);
}

function pageNameForName(pages = [], wanted = '') {
  const m = pages.find((p) => pageMatchesName(p, wanted));
  return first(get(m || {}, 'name'), get(m || {}, 'pageName'), wanted);
}

function responsibilityIdForPackage(pages = [], rows = [], packageName = '') {
  const page = pages.find((p) => pageMatchesName(p, packageName));
  const fromPage = first(get(page || {}, 'responsibilityId'), get(page || {}, 'meta.responsibilityId'));
  if (fromPage) return fromPage;
  const row = rows.find((r) => key(packageOf(r)) === key(packageName));
  return responsibilityIdOf(row || {});
}

function supplierForPackage(pages = [], rows = [], packageName = '', dependents = []) {
  const page = pages.find((p) => pageMatchesName(p, packageName));
  const fromPage = supplierOf(page || {});
  if (fromPage) return fromPage;
  const dep = dependents.find((d) => key(d.name) === key(packageName));
  if (dep?.supplier) return dep.supplier;
  const row = rows.find((r) => key(packageOf(r)) === key(packageName));
  return supplierOf(row || '');
}

function makeExactRowTargets(rows = []) {
  const out = [];
  const seen = new Set();
  for (const row of rows) {
    const rowId = rowIdOf(row);
    const pageId = pageIdOf(row);
    const pageName = pageNameOf(row);
    const packageName = packageOf(row);
    const block = blockOf(row);
    const level = levelOf(row);
    if (!rowId || !pageId || !packageName || !block || !level) continue;
    const dedupeKey = `${key(pageId)}|${key(rowId)}|${key(packageName)}|${key(block)}|${key(level)}`;
    if (seen.has(dedupeKey)) continue;
    seen.add(dedupeKey);
    out.push({ ...row, rowId, pageId, pageName, packageName, item: packageName, lineItemName: packageName, block, level, responsibilityId: responsibilityIdOf(row), supplier: supplierOf(row), uploadScopeKey: uploadScopeKeyOf(row) });
  }
  return out.sort((a, b) => numericAwareCompare(a.packageName, b.packageName) || compareByLevel(a, b));
}

function makeProjectBlockLevels(blockLevels = []) {
  const out = [];
  const seen = new Set();
  for (const item of blockLevels) {
    const block = first(item?.block, item?.name); const level = first(item?.level, item?.levelName);
    if (!block || !level) continue;
    const k = `${key(block)}|${key(level)}`;
    if (seen.has(k)) continue;
    seen.add(k);
    out.push({ ...item, block, level });
  }
  return out.sort(compareByLevel);
}

function uniqueFromRows(rows = [], getter) {
  const out = [];
  const seen = new Set();
  for (const row of rows) {
    const value = clean(getter(row));
    if (!value) continue;
    const k = key(value);
    if (seen.has(k)) continue;
    seen.add(k);
    out.push(value);
  }
  return out.sort(numericAwareCompare);
}

function OptionGroup({ title, items, value, onChange, emptyText }) {
  const cleanItems = (items || []).filter(Boolean);
  return (
    <View style={styles.card}>
      <Text style={styles.groupTitle}>{title}</Text>
      {cleanItems.length ? (
        <View style={styles.rowWrap}>
          {cleanItems.map((x) => (
            <Pill key={`${title}-${x}`} active={value === x} onPress={() => onChange(x)}>{x}</Pill>
          ))}
        </View>
      ) : (
        <Text style={styles.emptyText}>{emptyText}</Text>
      )}
    </View>
  );
}

function LocationCard({ icon, title, subtitle, active, onPress }) {
  return (
    <Pressable onPress={onPress} style={({ pressed }) => [styles.locationCard, active && styles.locationCardActive, pressed && styles.pressed]}>
      <View style={[styles.locationIcon, active && styles.locationIconActive]}>
        <Text style={[styles.locationIconText, active && styles.locationIconTextActive]}>{icon}</Text>
      </View>
      <View style={styles.locationTextBlock}>
        <Text style={styles.locationTitle}>{title}</Text>
        <Text style={styles.locationSubtitle}>{subtitle}</Text>
      </View>
      <Text style={[styles.chevron, active && styles.checkMark]}>{active ? '✓' : '›'}</Text>
    </Pressable>
  );
}

export default function UploadLocationScreen({ route, navigation }) {
  const { project } = route.params;
  const [apiTargets, setApiTargets] = useState(null);
  const [loadingTargets, setLoadingTargets] = useState(true);
  const [targetError, setTargetError] = useState('');

  // Derive schema from API response (with defaults)
  const schema = useMemo(() => ({ ...defaultSchema(), ...(apiTargets?.schema || {}) }), [apiTargets]);
  const targets = useMemo(() => extractTargets(project, apiTargets), [project, apiTargets]);

  const [scope, setScope] = useState('project');
  const [packageName, setPackageName] = useState('');
  const [projectBlock, setProjectBlock] = useState('');
  const [projectLevel, setProjectLevel] = useState('');
  const [packageBlock, setPackageBlock] = useState('');
  const [packageLevel, setPackageLevel] = useState('');
  const [selectedRow, setSelectedRow] = useState(null);
  const [uploadType, setUploadType] = useState('Images');

  const uploadTypes = ['Images', 'Drawings', 'Documents', 'Reports', 'Specifications', 'Instructions', 'QA'];
  const exactRowTargets = useMemo(() => makeExactRowTargets(targets.rows), [targets.rows]);
  const projectBlockLevels = useMemo(() => makeProjectBlockLevels(targets.blockLevels), [targets.blockLevels]);

  const projectBlocks = useMemo(() => uniqueFromRows(projectBlockLevels, (r) => r.block), [projectBlockLevels]);
  const projectLevelsForBlock = useMemo(() => uniqueFromRows(projectBlockLevels.filter((r) => key(r.block) === key(projectBlock)), (r) => r.level), [projectBlockLevels, projectBlock]);

  const rowsForPackage = useMemo(() => exactRowTargets.filter((r) => key(r.packageName) === key(packageName)), [exactRowTargets, packageName]);
  const packageBlocks = useMemo(() => uniqueFromRows(rowsForPackage, (r) => r.block), [rowsForPackage]);
  const rowsForPackageBlock = useMemo(() => rowsForPackage.filter((r) => key(r.block) === key(packageBlock)), [rowsForPackage, packageBlock]);
  const packageLevelsForBlock = useMemo(() => uniqueFromRows(rowsForPackageBlock, (r) => r.level), [rowsForPackageBlock]);
  const exactRowsForSelection = useMemo(() => rowsForPackageBlock.filter((r) => key(r.level) === key(packageLevel)), [rowsForPackageBlock, packageLevel]);

  const projectHomePageId = useMemo(() => pageIdForName(targets.pages || [], 'Project Home'), [targets.pages]);
  const projectHomePageName = useMemo(() => pageNameForName(targets.pages || [], 'Project Home') || 'Project Home', [targets.pages]);
  const packageHomePageId = useMemo(() => pageIdForName(targets.pages || [], packageName), [targets.pages, packageName]);
  const packageHomePageName = useMemo(() => pageNameForName(targets.pages || [], packageName) || packageName, [targets.pages, packageName]);
  const packageResponsibilityId = useMemo(() => responsibilityIdForPackage(targets.pages || [], exactRowTargets, packageName), [targets.pages, exactRowTargets, packageName]);
  const packageSupplierName = useMemo(() => supplierForPackage(targets.pages || [], exactRowTargets, packageName, targets.dependents || []), [targets.pages, exactRowTargets, packageName, targets.dependents]);

  useEffect(() => {
    let mounted = true;
    async function loadTargets() {
      setLoadingTargets(true);
      setTargetError('');
      try {
        const data = await fetchUploadTargets(project);
        if (mounted) setApiTargets(data);
      } catch (err) {
        if (mounted) setTargetError(err?.message || 'Could not load project metadata.');
      } finally {
        if (mounted) setLoadingTargets(false);
      }
    }
    loadTargets();
    return () => { mounted = false; };
  }, [project]);

  useEffect(() => {
    setPackageName((cur) => (targets.packages.includes(cur) ? cur : targets.packages[0] || ''));
  }, [targets.packages.join('|')]);

  useEffect(() => {
    if (scope !== 'project-block-level') return;
    setProjectBlock((cur) => (projectBlocks.includes(cur) ? cur : projectBlocks[0] || ''));
  }, [scope, projectBlocks.join('|')]);

  useEffect(() => {
    if (scope !== 'project-block-level') return;
    setProjectLevel((cur) => (projectLevelsForBlock.includes(cur) ? cur : projectLevelsForBlock[0] || ''));
  }, [scope, projectBlock, projectLevelsForBlock.join('|')]);

  useEffect(() => {
    if (scope !== 'package-block-level') return;
    setPackageBlock((cur) => (packageBlocks.includes(cur) ? cur : packageBlocks[0] || ''));
  }, [scope, packageName, packageBlocks.join('|')]);

  useEffect(() => {
    if (scope !== 'package-block-level') return;
    setPackageLevel((cur) => (packageLevelsForBlock.includes(cur) ? cur : packageLevelsForBlock[0] || ''));
  }, [scope, packageName, packageBlock, packageLevelsForBlock.join('|')]);

  useEffect(() => {
    if (scope !== 'package-block-level') return;
    setSelectedRow((cur) => {
      if (cur && exactRowsForSelection.some((r) => r.rowId === cur.rowId && r.pageId === cur.pageId)) return cur;
      return exactRowsForSelection[0] || null;
    });
  }, [scope, packageName, packageBlock, packageLevel, exactRowsForSelection]);

  const isProject = scope === 'project';
  const isProjectBlockLevel = scope === 'project-block-level';
  const isPackage = scope === 'package';
  const isPackageBlockLevel = scope === 'package-block-level';

  const canContinue = !loadingTargets && (
    isProject ||
    (isProjectBlockLevel && !!projectBlock && !!projectLevel) ||
    (isPackage && !!packageName) ||
    (isPackageBlockLevel && !!packageName && !!packageBlock && !!packageLevel && !!selectedRow?.rowId && !!selectedRow?.pageId)
  );

  const selectedRowSupplierName = useMemo(() => firstNonId(selectedRow?.supplier, packageSupplierName), [selectedRow, packageSupplierName]);

  const context = {
    project,
    scope,
    scopeType: isProject ? 'project' : isProjectBlockLevel ? 'master' : isPackage ? 'package' : 'row',
    packageName: isPackage || isPackageBlockLevel ? (isPackageBlockLevel ? selectedRow?.packageName || packageName : packageName) : '',
    block: isProjectBlockLevel ? projectBlock : isPackageBlockLevel ? selectedRow?.block || packageBlock : '',
    level: isProjectBlockLevel ? projectLevel : isPackageBlockLevel ? selectedRow?.level || packageLevel : '',
    uploadType,
    pageId: isPackageBlockLevel ? selectedRow?.pageId : isPackage ? packageHomePageId || packageName : projectHomePageId || 'Project Home',
    pageName: isPackageBlockLevel ? selectedRow?.pageName : isPackage ? packageHomePageName || packageName : projectHomePageName || 'Project Home',
    rowId: isPackageBlockLevel ? selectedRow?.rowId : '',
    item: isPackage || isPackageBlockLevel ? (isPackageBlockLevel ? selectedRow?.packageName || packageName : packageName) : 'Project Home',
    lineItemName: isPackage || isPackageBlockLevel ? (isPackageBlockLevel ? selectedRow?.packageName || packageName : packageName) : 'Project Home',
    responsibilityId: isPackageBlockLevel ? selectedRow?.responsibilityId || '' : isPackage ? packageResponsibilityId || '' : '',
    supplier: isPackageBlockLevel ? selectedRowSupplierName || '' : isPackage ? packageSupplierName || '' : 'Project Home',
    uploadScopeKey: isPackageBlockLevel ? selectedRow?.uploadScopeKey || '' : '',
    selectedRow: isPackageBlockLevel ? { ...selectedRow, supplier: selectedRowSupplierName || selectedRow?.supplier || '' } : null,
    schema,
  };

  const cL = schema.containerLabel;
  const scL = schema.subContainerLabel;
  const dL = schema.dependentLabel;

  const saveSummary = isPackageBlockLevel && selectedRow
    ? `${selectedRow.packageName} / ${selectedRow.block} / ${selectedRow.level}`
    : isPackage ? `${packageName || dL} / ${dL} wide`
    : isProjectBlockLevel ? `Project Home / ${projectBlock || cL} / ${projectLevel || scL}`
    : 'Project Home / Project wide';

  return (
    <Screen>
      <ScrollView showsVerticalScrollIndicator={false} contentContainerStyle={styles.scrollContent}>
        <Text style={styles.h1}>Choose upload location</Text>
        <Text style={styles.subText}>Select where this upload should be saved</Text>

        <View style={styles.summaryCard}>
          <Text style={styles.summaryLabel}>Will save to</Text>
          <View style={styles.summaryRow}>
            <View style={styles.summaryIcon}><Text style={styles.summaryIconText}>□</Text></View>
            <View style={styles.summaryTextBlock}>
              <Text style={styles.summaryTitle}>
                {isPackageBlockLevel ? `${dL} / ${cL} / ${scL}` : isPackage ? `${dL} Home` : isProjectBlockLevel ? `Project Home / ${cL} / ${scL}` : 'Project Home'}
              </Text>
              <Text style={styles.summarySub}>{saveSummary}</Text>
            </View>
          </View>
        </View>

        {loadingTargets ? (
          <View style={styles.loadingCard}>
            <ActivityIndicator />
            <Text style={styles.loadingText}>Loading project structure…</Text>
          </View>
        ) : null}

        {targetError ? <Text style={styles.errorText}>{targetError}</Text> : null}

        <Text style={styles.label}>Choose upload location</Text>
        <View style={styles.locationList}>
          <LocationCard icon="⌂" title="Project Home" subtitle="Save to Project Home" active={scope === 'project'} onPress={() => setScope('project')} />
          <LocationCard icon="▥" title={`Project Home / ${cL} / ${scL}`} subtitle={`Save to Project Home with ${cL.toLowerCase()} and ${scL.toLowerCase()}`} active={scope === 'project-block-level'} onPress={() => setScope('project-block-level')} />
          <LocationCard icon="▱" title={`${dL} Home`} subtitle={`Save to ${dL.toLowerCase()} home`} active={scope === 'package'} onPress={() => setScope('package')} />
          <LocationCard icon="☷" title={`${dL} / ${cL} / ${scL}`} subtitle={`Save to specific ${dL.toLowerCase()}, ${cL.toLowerCase()} and ${scL.toLowerCase()}`} active={scope === 'package-block-level'} onPress={() => setScope('package-block-level')} />
        </View>

        <Text style={styles.label}>Upload type</Text>
        <View style={styles.gridWrap}>
          {uploadTypes.map((t) => (
            <Pill key={t} active={uploadType === t} onPress={() => setUploadType(t)} style={styles.gridPill}>{t}</Pill>
          ))}
        </View>

        {isProjectBlockLevel ? (
          <View style={styles.selectorBlock}>
            <OptionGroup title={`1. ${cL}`} items={projectBlocks} value={projectBlock}
              onChange={(value) => {
                const nextLevels = uniqueFromRows(projectBlockLevels.filter((r) => key(r.block) === key(value)), (r) => r.level);
                setProjectBlock(value); setProjectLevel(nextLevels[0] || '');
              }}
              emptyText={`No ${schema.containerPlural.toLowerCase()} are configured in Project Home.`} />
            <OptionGroup title={`2. ${scL}`} items={projectLevelsForBlock} value={projectLevel} onChange={setProjectLevel}
              emptyText={`No ${schema.subContainerPlural.toLowerCase()} are available for the selected ${cL.toLowerCase()}.`} />
          </View>
        ) : null}

        {isPackage ? (
          <OptionGroup title={dL} items={targets.packages} value={packageName} onChange={setPackageName}
            emptyText={`No ${schema.dependentPlural.toLowerCase()} are configured for this project.`} />
        ) : null}

        {isPackageBlockLevel ? (
          <View style={styles.selectorBlock}>
            <OptionGroup title={`1. ${dL}`} items={targets.packages} value={packageName}
              onChange={(value) => {
                const nextRows = exactRowTargets.filter((r) => key(r.packageName) === key(value));
                const nextBlocks = uniqueFromRows(nextRows, (r) => r.block);
                const nextBlock = nextBlocks[0] || '';
                const nextRowsForBlock = nextRows.filter((r) => key(r.block) === key(nextBlock));
                const nextLevels = uniqueFromRows(nextRowsForBlock, (r) => r.level);
                const nextLevel = nextLevels[0] || '';
                const nextSelection = nextRowsForBlock.find((r) => key(r.level) === key(nextLevel)) || null;
                setPackageName(value); setPackageBlock(nextBlock); setPackageLevel(nextLevel); setSelectedRow(nextSelection);
              }}
              emptyText={`No ${schema.dependentPlural.toLowerCase()} are configured for this project.`} />
            <OptionGroup title={`2. ${cL}`} items={packageBlocks} value={packageBlock}
              onChange={(value) => {
                const nextRowsForBlock = rowsForPackage.filter((r) => key(r.block) === key(value));
                const nextLevels = uniqueFromRows(nextRowsForBlock, (r) => r.level);
                const nextLevel = nextLevels[0] || '';
                const nextSelection = nextRowsForBlock.find((r) => key(r.level) === key(nextLevel)) || null;
                setPackageBlock(value); setPackageLevel(nextLevel); setSelectedRow(nextSelection);
              }}
              emptyText={`No ${schema.containerPlural.toLowerCase()} are available for the selected ${dL.toLowerCase()}.`} />
            <OptionGroup title={`3. ${scL}`} items={packageLevelsForBlock} value={packageLevel}
              onChange={(value) => {
                const nextSelection = rowsForPackageBlock.find((r) => key(r.level) === key(value)) || null;
                setPackageLevel(value); setSelectedRow(nextSelection);
              }}
              emptyText={`No ${schema.subContainerPlural.toLowerCase()} are available for the selected ${dL.toLowerCase()} and ${cL.toLowerCase()}.`} />
          </View>
        ) : null}

        {!canContinue && !loadingTargets ? (
          <Text style={styles.warningText}>This upload location needs the required web-app metadata before you can continue.</Text>
        ) : null}

        <View style={styles.footerAction}>
          <AppButton title="Next" disabled={!canContinue} onPress={() => navigation.navigate('Upload', { context })} />
        </View>
      </ScrollView>
    </Screen>
  );
}

const styles = StyleSheet.create({
  scrollContent: { paddingBottom: 34 },
  h1: { color: colors.dark, fontSize: 27, fontWeight: '900', marginBottom: 8, letterSpacing: -0.6 },
  subText: { color: colors.muted, fontSize: 14, marginBottom: 20 },
  label: { fontSize: 13, fontWeight: '700', color: colors.dark, marginBottom: 8, marginTop: 16, textTransform: 'uppercase', letterSpacing: 0.5 },
  summaryCard: { backgroundColor: colors.surface, borderRadius: 14, padding: 16, marginBottom: 20, borderWidth: 1, borderColor: colors.border },
  summaryLabel: { fontSize: 11, color: colors.muted, fontWeight: '600', marginBottom: 10, textTransform: 'uppercase', letterSpacing: 0.5 },
  summaryRow: { flexDirection: 'row', alignItems: 'center', gap: 12 },
  summaryIcon: { width: 36, height: 36, borderRadius: 8, backgroundColor: colors.primary + '15', alignItems: 'center', justifyContent: 'center' },
  summaryIconText: { fontSize: 18 },
  summaryTextBlock: { flex: 1 },
  summaryTitle: { fontSize: 15, fontWeight: '700', color: colors.dark },
  summarySub: { fontSize: 13, color: colors.muted, marginTop: 2 },
  loadingCard: { flexDirection: 'row', alignItems: 'center', gap: 10, padding: 14, backgroundColor: colors.surface, borderRadius: 12, marginBottom: 12, borderWidth: 1, borderColor: colors.border },
  loadingText: { color: colors.muted, fontSize: 13 },
  errorText: { color: colors.danger || '#EF4444', fontSize: 13, marginBottom: 12 },
  warningText: { color: colors.warning || '#F59E0B', fontSize: 13, marginBottom: 12 },
  locationList: { gap: 8, marginBottom: 8 },
  locationCard: { flexDirection: 'row', alignItems: 'center', gap: 12, padding: 14, backgroundColor: colors.surface, borderRadius: 14, borderWidth: 1.5, borderColor: colors.border },
  locationCardActive: { borderColor: colors.primary, backgroundColor: colors.primary + '08' },
  pressed: { opacity: 0.75 },
  locationIcon: { width: 40, height: 40, borderRadius: 10, backgroundColor: colors.border, alignItems: 'center', justifyContent: 'center' },
  locationIconActive: { backgroundColor: colors.primary + '20' },
  locationIconText: { fontSize: 18, color: colors.muted },
  locationIconTextActive: { color: colors.primary },
  locationTextBlock: { flex: 1 },
  locationTitle: { fontSize: 15, fontWeight: '700', color: colors.dark },
  locationSubtitle: { fontSize: 12, color: colors.muted, marginTop: 1 },
  chevron: { fontSize: 18, color: colors.muted },
  checkMark: { color: colors.primary, fontWeight: '700' },
  gridWrap: { flexDirection: 'row', flexWrap: 'wrap', gap: 8, marginBottom: 4 },
  gridPill: { minWidth: 90 },
  selectorBlock: { gap: 12 },
  card: { backgroundColor: colors.surface, borderRadius: 14, padding: 14, borderWidth: 1, borderColor: colors.border },
  groupTitle: { fontSize: 13, fontWeight: '700', color: colors.dark, marginBottom: 10, textTransform: 'uppercase', letterSpacing: 0.5 },
  rowWrap: { flexDirection: 'row', flexWrap: 'wrap', gap: 8 },
  emptyText: { fontSize: 13, color: colors.muted, fontStyle: 'italic' },
  footerAction: { marginTop: 24 },
});
