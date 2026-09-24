/**
 * Patch @univerjs/ui ContextMenu so only one sibling submenu is open at a time.
 * Upstream keeps per-item submenuVisible + 500ms delayed close, so hovering
 * 插入→删除→冻结→保护 leaves stacked portals (and leftover portals cancel their
 * own close timers when the pointer grazes them).
 *
 * Idempotent via GenOffice markers. Run after npm install / before sheets build:
 *   node scripts/patch-univer-ui-context-menu.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const MARKER = 'GenOffice: exclusive-context-menu-submenu';

const targets = [
  path.join(root, 'node_modules/@univerjs/ui/lib/es/index.js'),
  path.join(root, 'node_modules/@univerjs/ui/lib/cjs/index.js'),
  path.join(root, 'node_modules/@univerjs/ui/lib/index.js'),
];

const INSERT_CTX = `
/* ${MARKER} */
const ContextMenuSubmenuExclusiveContext = createContext(null);
`;

function patchFile(file) {
  if (!fs.existsSync(file)) {
    console.warn('[patch-univer-ui-context-menu] skip missing', path.relative(root, file));
    return false;
  }
  let src = fs.readFileSync(file, 'utf8');
  if (src.includes(MARKER)) {
    console.log('[patch-univer-ui-context-menu] already patched', path.relative(root, file));
    return false;
  }

  const anchor = 'const CONTEXT_MENU_SUBMENU_PORTAL_ATTR = "data-u-context-menu-submenu";';
  if (!src.includes(anchor)) {
    console.warn('[patch-univer-ui-context-menu] anchor missing', path.relative(root, file));
    return false;
  }
  src = src.replace(anchor, `${anchor}${INSERT_CTX}`);

  // Wrap ContextMenuMenu body with exclusive provider.
  const menuStart = 'function ContextMenuMenu(props) {\n\tconst { menuSchemas, menuSessionVersion, submenuPortalContainer, activeItemIds, hiddenItemIds, onOptionSelect, maxMenuHeight } = props;\n\tconst localeService = useDependency(LocaleService);\n\tconst hiddenGroupStates = useContextGroupHiddenStates$1(menuSchemas);\n\tconst visibleSchemas = useMemo(() => {\n\t\treturn menuSchemas.filter((item) => {\n\t\t\tif (!hasRenderableContextMenuSchema(item)) return false;\n\t\t\tif (!item.children) return true;\n\t\t\treturn !hiddenGroupStates[item.key];\n\t\t});\n\t}, [hiddenGroupStates, menuSchemas]);\n\treturn /* @__PURE__ */ jsx(Fragment$1, { children: visibleSchemas.map((menuSchema, index) => {';

  const menuStartPatched = `function ContextMenuMenu(props) {
\tconst { menuSchemas, menuSessionVersion, submenuPortalContainer, activeItemIds, hiddenItemIds, onOptionSelect, maxMenuHeight } = props;
\tconst localeService = useDependency(LocaleService);
\tconst hiddenGroupStates = useContextGroupHiddenStates$1(menuSchemas);
\tconst visibleSchemas = useMemo(() => {
\t\treturn menuSchemas.filter((item) => {
\t\t\tif (!hasRenderableContextMenuSchema(item)) return false;
\t\t\tif (!item.children) return true;
\t\t\treturn !hiddenGroupStates[item.key];
\t\t});
\t}, [hiddenGroupStates, menuSchemas]);
\t/* ${MARKER} */
\tconst [activeSubmenuKey, setActiveSubmenuKey] = useState(null);
\tconst exclusiveApi = useMemo(() => ({
\t\tactiveSubmenuKey,
\t\topenSubmenu: (key) => setActiveSubmenuKey(key),
\t\tcloseSubmenu: (key) => setActiveSubmenuKey((cur) => cur === key ? null : cur),
\t\tcloseAll: () => setActiveSubmenuKey(null),
\t}), [activeSubmenuKey]);
\treturn /* @__PURE__ */ jsx(ContextMenuSubmenuExclusiveContext.Provider, { value: exclusiveApi, children: /* @__PURE__ */ jsx(Fragment$1, { children: visibleSchemas.map((menuSchema, index) => {`;

  if (!src.includes(menuStart)) {
    console.warn('[patch-univer-ui-context-menu] ContextMenuMenu start missing', path.relative(root, file));
    return false;
  }
  src = src.replace(menuStart, menuStartPatched);

  // Close the extra Provider wrapper: `}) });` before ContextMenuMenuItem → `}) }) });`
  const menuEnd = `\t}) });
}
function ContextMenuMenuItem(props) {`;
  const menuEndPatched = `\t}) }) });
}
function ContextMenuMenuItem(props) {`;
  if (!src.includes(menuEnd)) {
    console.warn('[patch-univer-ui-context-menu] ContextMenuMenu end missing', path.relative(root, file));
    return false;
  }
  src = src.replace(menuEnd, menuEndPatched);

  // Item: use exclusive open key instead of sticky local visibility.
  const itemState = `\tconst [submenuVisible, setSubmenuVisible] = useState(false);
\tconst [submenuPosition, setSubmenuPosition] = useState({
\t\tleft: 0,
\t\ttop: 0
\t});`;

  const itemStatePatched = `\t/* ${MARKER} */
\tconst exclusive = useContext(ContextMenuSubmenuExclusiveContext);
\tconst submenuVisible = !!(exclusive && exclusive.activeSubmenuKey === menuKey);
\tconst setSubmenuVisible = useCallback((next) => {
\t\tif (!exclusive) return;
\t\tif (next) exclusive.openSubmenu(menuKey);
\t\telse exclusive.closeSubmenu(menuKey);
\t}, [exclusive, menuKey]);
\tconst [submenuPosition, setSubmenuPosition] = useState({
\t\tleft: 0,
\t\ttop: 0
\t});`;

  if (!src.includes(itemState)) {
    console.warn('[patch-univer-ui-context-menu] submenuVisible state missing', path.relative(root, file));
    return false;
  }
  src = src.replace(itemState, itemStatePatched);

  const mouseEnter = `\t\tonMouseEnter: () => {
\t\t\tclearSubmenuCloseTimer();
\t\t\tif (hasSubmenu && !disabled) {
\t\t\t\tsetSubmenuPositionReady(false);
\t\t\t\tsetSubmenuVisible(true);
\t\t\t}
\t\t},`;

  const mouseEnterPatched = `\t\tonMouseEnter: () => {
\t\t\tclearSubmenuCloseTimer();
\t\t\tif (hasSubmenu && !disabled) {
\t\t\t\tsetSubmenuPositionReady(false);
\t\t\t\tsetSubmenuVisible(true);
\t\t\t} else if (exclusive) {
\t\t\t\texclusive.closeAll();
\t\t\t}
\t\t},`;

  if (!src.includes(mouseEnter)) {
    console.warn('[patch-univer-ui-context-menu] onMouseEnter missing', path.relative(root, file));
    return false;
  }
  // Only patch the ContextMenuMenuItem mouseEnter (first occurrence after our marker in item).
  // There may be similar patterns elsewhere — replace only within ContextMenuMenuItem by
  // limiting to the first match after "function ContextMenuMenuItem".
  const itemIdx = src.indexOf('function ContextMenuMenuItem(props)');
  const before = src.slice(0, itemIdx);
  const after = src.slice(itemIdx);
  if (!after.includes(mouseEnter)) {
    console.warn('[patch-univer-ui-context-menu] item onMouseEnter missing', path.relative(root, file));
    return false;
  }
  src = before + after.replace(mouseEnter, mouseEnterPatched);

  fs.writeFileSync(file, src);
  console.log('[patch-univer-ui-context-menu] patched', path.relative(root, file));
  return true;
}

let n = 0;
for (const t of targets) {
  if (patchFile(t)) n++;
}
console.log(`[patch-univer-ui-context-menu] done (${n} files updated).`);
