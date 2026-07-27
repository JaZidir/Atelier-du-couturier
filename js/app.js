/**
 * app.js
 * ---------------------------------------------------------------------------
 * Point d'entrée et logique applicative de "Atelier du Couturier".
 *
 * Architecture (volontairement modulaire pour permettre les évolutions
 * futures listées dans le cahier des charges : stock, historique,
 * fabrication, Electron, ré-import d'un registre) :
 *
 *   - ExcelReader (excel.js) : lecture pure du fichier .xlsx -> données brutes
 *   - AppState                : source de vérité en mémoire
 *   - Recompute engine        : dérive coût / bénéfice / rentabilité
 *   - Render layer            : fonctions render*() qui projettent l'état vers le DOM
 *   - Event handlers          : traduisent les interactions utilisateur en mutations d'état
 *
 * Aucune donnée n'est codée en dur ici : tout provient du registre Excel.
 * ---------------------------------------------------------------------------
 */

// =============================================================================
// ÉTAT DE L'APPLICATION
// =============================================================================

const AppState = {
  materiaux: [],          // [{ id, name, buyPrice, quantity, unitPrice, originalBuyPrice, customBuyPrice }]
  objets: [],             // [{ id, name, category, sellPrice, recipe, excelCost, ... }]
  materiauxById: new Map(),
  categories: ['Tous'],
  activeTab: 'catalogue',  // 'catalogue' | 'rachat'
  selectedItemId: null,
  selectedCategory: 'Tous',
  searchQuery: '',
  isLoaded: false,
  // Table de correspondance optionnelle "imageKey -> URL ou chemin d'image",
  // chargée depuis data/images.json si le fichier existe (voir loadImageMap()).
  // Permet d'associer des images (ex. hébergées ailleurs) sans toucher au code.
  imageMap: {},
  // Table de variantes visuelles (couleurs multiples par objet), même
  // principe que imageMap : IMAGE_MAP embarquée + data/variants.json en
  // complément (voir loadVariantsMap()).
  variantsMap: {},
  // Index de la variante d'image actuellement affichée pour l'objet
  // sélectionné (voir VARIANTS_MAP dans js/variants-map.js). Remis à 0
  // à chaque changement d'objet sélectionné.
  selectedVariantIndex: 0
};

// Chemin de l'image par défaut affichée quand aucune image n'est trouvée pour un objet.
const DEFAULT_IMAGE = 'images/placeholder.svg';

// =============================================================================
// MOTEUR DE RECALCUL
// =============================================================================

/**
 * Prix de rachat courant d'un matériau : la surcharge utilisateur si elle
 * existe (onglet "Rachat des matériaux"), sinon la valeur du registre Excel.
 */
function getCurrentBuyPrice(material) {
  return material.customBuyPrice !== null && material.customBuyPrice !== undefined
    ? material.customBuyPrice
    : material.originalBuyPrice;
}

/** Prix unitaire courant d'un matériau, recalculé à partir du prix de rachat courant. */
function getCurrentUnitPrice(material) {
  const buyPrice = getCurrentBuyPrice(material);
  return material.quantity > 0 ? buyPrice / material.quantity : buyPrice;
}

/**
 * Calcule le coût de fabrication réel d'un objet à partir de sa recette
 * et des prix de matériaux actuels (donc toujours à jour après une
 * modification dans l'onglet "Rachat des matériaux").
 */
function computeItemCost(item) {
  let cost = 0;
  const missing = [];

  for (const line of item.recipe) {
    const material = AppState.materiauxById.get(line.materialId);
    if (!material) {
      missing.push(line.materialName);
      continue;
    }
    cost += getCurrentUnitPrice(material) * line.quantity;
  }

  return { cost, missing };
}

/** Calcule coût, bénéfice et rentabilité (%) d'un objet. */
function computeItemStats(item) {
  const { cost, missing } = computeItemCost(item);
  const profit = item.sellPrice - cost;
  const rentabilite = cost > 0 ? profit / cost : (item.sellPrice > 0 ? 1 : 0);
  return { cost, profit, rentabilite, missing };
}

// =============================================================================
// CHARGEMENT DES DONNÉES
// =============================================================================

async function init() {
  bindStaticEvents();
  await loadImageMap();
  await loadVariantsMap();
  try {
    const data = await ExcelReader.loadFromDefaultPath('data/registre.xlsx');
    setData(data);
  } catch (err) {
    // Aucun registre embarqué : on invite l'utilisateur à en importer un.
    showEmptyState();
  }
}

/**
 * Initialise la table de correspondance d'images.
 * Base : IMAGE_MAP embarquée dans js/images-map.js (toujours disponible,
 * y compris en ouvrant index.html directement en double-clic).
 * Complément optionnel : data/images.json, si l'application tourne via un
 * serveur local (http/https) où fetch() sur un fichier local fonctionne.
 * Ses entrées prennent le pas sur IMAGE_MAP en cas de conflit, pour que
 * modifier data/images.json reste le moyen normal de mettre à jour les
 * images sans toucher au code.
 */
async function loadImageMap() {
  AppState.imageMap = { ...(typeof IMAGE_MAP !== 'undefined' ? IMAGE_MAP : {}) };
  try {
    const response = await fetch('data/images.json');
    if (!response.ok) return;
    const fetched = await response.json();
    AppState.imageMap = { ...AppState.imageMap, ...fetched };
  } catch (err) {
    // fetch() indisponible (ex: ouverture en file://) : on garde IMAGE_MAP tel quel.
  }
}

/**
 * Même principe que loadImageMap(), mais pour les variantes visuelles
 * (couleurs multiples par objet, voir js/variants-map.js). data/variants.json,
 * quand fetch() fonctionne, complète/écrase VARIANTS_MAP — c'est le fichier
 * à modifier en priorité car il ne nécessite pas de savoir lire du code.
 */
async function loadVariantsMap() {
  AppState.variantsMap = { ...(typeof VARIANTS_MAP !== 'undefined' ? VARIANTS_MAP : {}) };
  try {
    const response = await fetch('data/variants.json');
    if (!response.ok) return;
    const fetched = await response.json();
    AppState.variantsMap = { ...AppState.variantsMap, ...fetched };
  } catch (err) {
    // fetch() indisponible : on garde VARIANTS_MAP tel quel.
  }
}

function setData({ materiaux, objets }) {
  AppState.materiaux = materiaux;
  AppState.objets = objets;
  AppState.materiauxById = new Map(materiaux.map(m => [m.id, m]));
  AppState.categories = ['Tous', ...Array.from(new Set(objets.map(o => o.category))).sort()];
  AppState.isLoaded = true;
  AppState.selectedItemId = objets.length > 0 ? objets[0].id : null;
  AppState.selectedCategory = 'Tous';
  AppState.searchQuery = '';

  hideEmptyState();
  renderAll();
  showToast(`Registre chargé : ${objets.length} objets, ${materiaux.length} matériaux.`, 'success');
}

async function handleFileInput(file) {
  try {
    const buffer = await file.arrayBuffer();
    const data = ExcelReader.loadFromArrayBuffer(buffer);
    setData(data);
  } catch (err) {
    showToast(`Impossible de lire ce fichier : ${err.message}`, 'error');
  }
}

// =============================================================================
// RENDU — ÉTAT VIDE (aucun registre chargé)
// =============================================================================

function showEmptyState() {
  document.getElementById('empty-state').classList.remove('hidden');
  document.getElementById('app-content').classList.add('hidden');
}

function hideEmptyState() {
  document.getElementById('empty-state').classList.add('hidden');
  document.getElementById('app-content').classList.remove('hidden');
}

// =============================================================================
// RENDU GLOBAL
// =============================================================================

function renderAll() {
  renderCategoryList();
  renderItemList();
  renderDetailPanel();
  renderRachatTable();
  renderTabs();
}

function renderTabs() {
  document.querySelectorAll('.tab-button').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.tab === AppState.activeTab);
  });
  document.getElementById('panel-catalogue').classList.toggle('hidden', AppState.activeTab !== 'catalogue');
  document.getElementById('panel-rachat').classList.toggle('hidden', AppState.activeTab !== 'rachat');
}

// ---------------------------------------------------------------------------
// Onglet Catalogue — liste des catégories
// ---------------------------------------------------------------------------

function renderCategoryList() {
  const container = document.getElementById('category-list');
  container.innerHTML = '';

  AppState.categories.forEach(cat => {
    const count = cat === 'Tous'
      ? AppState.objets.length
      : AppState.objets.filter(o => o.category === cat).length;

    const btn = document.createElement('button');
    btn.className = 'category-item' + (cat === AppState.selectedCategory ? ' active' : '');
    btn.innerHTML = `<span>${escapeHtml(cat)}</span><span class="category-count">${count}</span>`;
    btn.addEventListener('click', () => {
      AppState.selectedCategory = cat;
      renderCategoryList();
      renderItemList();
    });
    container.appendChild(btn);
  });
}

// ---------------------------------------------------------------------------
// Onglet Catalogue — liste des objets (recherche + filtre catégorie)
// ---------------------------------------------------------------------------

function getFilteredItems() {
  const query = ExcelReader.utils.normalize(AppState.searchQuery);
  return AppState.objets.filter(item => {
    const matchesCategory = AppState.selectedCategory === 'Tous' || item.category === AppState.selectedCategory;
    const matchesQuery = !query || ExcelReader.utils.normalize(item.name).includes(query);
    return matchesCategory && matchesQuery;
  });
}

function renderItemList() {
  const container = document.getElementById('item-list');
  container.innerHTML = '';

  const items = getFilteredItems();

  if (items.length === 0) {
    container.innerHTML = '<p class="list-empty">Aucun objet ne correspond à cette recherche.</p>';
    return;
  }

  items.forEach(item => {
    const { rentabilite } = computeItemStats(item);
    const row = document.createElement('button');
    row.className = 'item-row' + (item.id === AppState.selectedItemId ? ' active' : '');
    row.innerHTML = `
      <span class="item-row-name">${escapeHtml(item.name)}</span>
      <span class="rentability-badge ${rentabilityClass(rentabilite)}">${formatPercent(rentabilite)}</span>
    `;
    row.addEventListener('click', () => {
      AppState.selectedItemId = item.id;
      AppState.selectedVariantIndex = 0;
      renderItemList();
      renderDetailPanel();
    });
    container.appendChild(row);
  });
}

// ---------------------------------------------------------------------------
// Onglet Catalogue — panneau de détail (droite)
// ---------------------------------------------------------------------------

function renderDetailPanel() {
  const panel = document.getElementById('detail-panel');
  const item = AppState.objets.find(o => o.id === AppState.selectedItemId);

  if (!item) {
    panel.innerHTML = '<p class="detail-empty">Sélectionnez un objet dans la liste pour voir ses détails.</p>';
    return;
  }

  const { cost, profit, rentabilite, missing } = computeItemStats(item);

  const recipeHtml = item.recipe.length
    ? item.recipe.map(line => {
        const material = AppState.materiauxById.get(line.materialId);
        const isMissing = !material;
        return `
          <li class="recipe-line ${isMissing ? 'missing' : ''}">
            <span class="recipe-name">${escapeHtml(line.materialName)}</span>
            <span class="recipe-qty">x${line.quantity}</span>
          </li>`;
      }).join('')
    : '<li class="recipe-line missing">Aucune recette renseignée dans le registre.</li>';

  const warningHtml = missing.length
    ? `<div class="detail-warning">⚠ Matériau introuvable dans la feuille MATERIAUX : ${missing.map(escapeHtml).join(', ')}</div>`
    : '';

  const variants = getVariantsForItem(item.imageKey);
  const variantIndex = Math.min(AppState.selectedVariantIndex, variants.length - 1);
  const variant = variants[variantIndex];
  const hasMultipleVariants = variants.length > 1;

  const variantSelectHtml = hasMultipleVariants
    ? `
      <div class="variant-picker">
        <label for="variant-select">Variante</label>
        <select id="variant-select">
          ${variants.map((v, i) => `<option value="${i}" ${i === variantIndex ? 'selected' : ''}>${escapeHtml(v.label || item.name)}</option>`).join('')}
        </select>
      </div>`
    : '';

  panel.innerHTML = `
    <div class="detail-images-row">
      <div class="detail-image-wrap">
        <img class="detail-image" alt="${escapeHtml(item.name)} - porté homme" ${imageAttrs(variant, item.imageKey, 'homme')}>
        <span class="detail-image-label">Porté homme</span>
      </div>
      <div class="detail-image-wrap">
        <img class="detail-image" alt="${escapeHtml(item.name)} - porté femme" ${imageAttrs(variant, item.imageKey, 'femme')}>
        <span class="detail-image-label">Porté femme</span>
      </div>
    </div>

    ${variantSelectHtml}

    <div class="detail-header">
      <span class="detail-category">${escapeHtml(item.category)}</span>
      <h2 class="detail-title">${escapeHtml(item.name)}</h2>
    </div>

    ${warningHtml}

    <div class="stats-grid">
      <div class="stat-card">
        <span class="stat-label">Coût de fabrication</span>
        <span class="stat-value">${formatNumber(cost)}</span>
      </div>
      <div class="stat-card">
        <span class="stat-label">Prix de vente</span>
        <span class="stat-value">${formatNumber(item.sellPrice)}</span>
      </div>
      <div class="stat-card">
        <span class="stat-label">Bénéfice</span>
        <span class="stat-value ${profit >= 0 ? 'positive' : 'negative'}">${formatNumber(profit)}</span>
      </div>
      <div class="stat-card">
        <span class="stat-label">Rentabilité</span>
        <span class="stat-value ${rentabilityClass(rentabilite)}">${formatPercent(rentabilite)}</span>
      </div>
    </div>

    <div class="recipe-section">
      <h3 class="section-label">Recette</h3>
      <ul class="recipe-list">${recipeHtml}</ul>
    </div>

    <button class="btn-primary btn-fabriquer" id="btn-fabriquer">Fabriquer</button>
  `;

  if (hasMultipleVariants) {
    document.getElementById('variant-select').addEventListener('change', evt => {
      AppState.selectedVariantIndex = parseInt(evt.target.value, 10);
      renderDetailPanel();
    });
  }


  document.getElementById('btn-fabriquer').addEventListener('click', () => {
    // La gestion du stock / fabrication réelle est prévue dans une évolution
    // future (voir README). Pour l'instant on informe simplement l'atelier.
    showToast('Gestion de la fabrication à venir dans une prochaine version.', 'info');
  });
}

// ---------------------------------------------------------------------------
// Onglet Rachat des matériaux
// ---------------------------------------------------------------------------

function renderRachatTable() {
  const tbody = document.getElementById('rachat-tbody');
  tbody.innerHTML = '';

  AppState.materiaux
    .slice()
    .sort((a, b) => a.name.localeCompare(b.name, 'fr'))
    .forEach(material => {
      const currentBuyPrice = getCurrentBuyPrice(material);
      const currentUnitPrice = getCurrentUnitPrice(material);
      const isOverridden = material.customBuyPrice !== null && material.customBuyPrice !== undefined;

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${escapeHtml(material.name)}</td>
        <td>
          <input
            type="number"
            class="rachat-input"
            min="0"
            step="0.01"
            value="${currentBuyPrice}"
            data-material-id="${material.id}"
          >
        </td>
        <td class="muted">${material.quantity}</td>
        <td class="muted">${formatNumber(currentUnitPrice)}</td>
        <td>
          <button class="btn-reset" data-material-id="${material.id}" ${isOverridden ? '' : 'disabled'}>
            Réinitialiser
          </button>
        </td>
      `;
      tbody.appendChild(tr);
    });

  tbody.querySelectorAll('.rachat-input').forEach(input => {
    input.addEventListener('input', onRachatInputChange);
  });
  tbody.querySelectorAll('.btn-reset').forEach(btn => {
    btn.addEventListener('click', onRachatReset);
  });
}

function onRachatInputChange(evt) {
  const materialId = evt.target.dataset.materialId;
  const material = AppState.materiauxById.get(materialId);
  if (!material) return;

  const value = parseFloat(evt.target.value);
  material.customBuyPrice = isNaN(value) ? null : value;

  // Le changement de prix impacte potentiellement tous les objets du
  // catalogue : on recalcule et on rafraîchit les deux onglets.
  renderItemList();
  renderDetailPanel();
  renderRachatTable();
}

function onRachatReset(evt) {
  const materialId = evt.target.dataset.materialId;
  const material = AppState.materiauxById.get(materialId);
  if (!material) return;

  material.customBuyPrice = null;
  renderItemList();
  renderDetailPanel();
  renderRachatTable();
}

// =============================================================================
// RÉSOLUTION DES IMAGES
// =============================================================================

/**
 * Retourne la liste des variantes visuelles d'un objet : chacune a un
 * libellé (affiché dans le menu déroulant) et ses propres fichiers
 * homme/femme. Source : VARIANTS_MAP (js/variants-map.js), pour les objets
 * qui ont plusieurs couleurs/déclinaisons non présentes dans le registre
 * Excel. Si aucune variante n'est définie, on retourne une variante
 * implicite unique basée sur la convention habituelle
 * (images/<imageKey>-homme.png / -femme.png) — le menu déroulant n'est
 * alors pas affiché.
 */
function getVariantsForItem(imageKey) {
  const defined = AppState.variantsMap[imageKey] || null;
  if (defined && defined.length > 0) return defined;
  return [{ label: null, homme: `images/${imageKey}-homme.png`, femme: `images/${imageKey}-femme.png` }];
}

/**
 * Construit la liste ordonnée des sources à essayer pour une image, du
 * plus fiable au moins fiable. handleImageError() avance dans cette liste
 * à chaque échec de chargement.
 *
 * @param {object} variant la variante actuellement sélectionnée ({label, homme, femme})
 * @param {string} imageKey identifiant de l'objet (pour le repli sur l'URL externe)
 * @param {string|null} gender 'homme' | 'femme' | null (null = vignette
 *   générique de la liste, qui n'a pas de déclinaison par genre)
 */
function buildImageCandidates(variant, imageKey, gender) {
  const candidates = [];
  if (gender) {
    candidates.push(variant[gender]);
  } else {
    candidates.push(variant.homme);
    candidates.push(variant.femme);
  }
  const mapped = AppState.imageMap[imageKey];
  if (mapped) candidates.push(mapped);
  candidates.push(DEFAULT_IMAGE);
  return candidates;
}

/** Attributs HTML (src + données de cascade) à injecter dans une balise <img>. */
function imageAttrs(variant, imageKey, gender) {
  const candidates = buildImageCandidates(variant, imageKey, gender);
  const encoded = escapeHtml(JSON.stringify(candidates));
  return `src="${candidates[0]}" data-candidates="${encoded}" data-candidate-index="0" referrerpolicy="no-referrer" onerror="handleImageError(this)"`;
}

/**
 * Gère l'échec de chargement d'une image en passant au candidat suivant
 * de la cascade construite par buildImageCandidates() : fichier local
 * (homme/femme) -> URL externe connue -> images/placeholder.svg.
 */
function handleImageError(imgEl) {
  const candidates = JSON.parse(imgEl.dataset.candidates || '[]');
  const nextIndex = parseInt(imgEl.dataset.candidateIndex || '0', 10) + 1;
  if (nextIndex < candidates.length) {
    imgEl.dataset.candidateIndex = String(nextIndex);
    imgEl.src = candidates[nextIndex];
  }
  // Sinon : on est déjà sur le placeholder, qui ne peut pas échouer davantage.
}

// =============================================================================
// FORMATAGE
// =============================================================================

function formatNumber(value) {
  return new Intl.NumberFormat('fr-FR', { maximumFractionDigits: 2 }).format(value || 0);
}

function formatPercent(value) {
  return new Intl.NumberFormat('fr-FR', { style: 'percent', maximumFractionDigits: 0 }).format(value || 0);
}

function rentabilityClass(rentabilite) {
  if (rentabilite >= 0.5) return 'excellent';
  if (rentabilite >= 0.15) return 'good';
  if (rentabilite >= 0) return 'low';
  return 'negative';
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str ?? '';
  return div.innerHTML;
}

// =============================================================================
// NOTIFICATIONS (toasts)
// =============================================================================

let toastTimeoutId = null;

function showToast(message, type = 'info') {
  const toast = document.getElementById('toast');
  toast.textContent = message;
  toast.className = `toast toast-${type} visible`;

  clearTimeout(toastTimeoutId);
  toastTimeoutId = setTimeout(() => {
    toast.classList.remove('visible');
  }, 3500);
}

// =============================================================================
// ÉVÉNEMENTS STATIQUES (liés une seule fois, au chargement de la page)
// =============================================================================

function bindStaticEvents() {
  // Onglets principaux
  document.querySelectorAll('.tab-button').forEach(btn => {
    btn.addEventListener('click', () => {
      AppState.activeTab = btn.dataset.tab;
      renderTabs();
    });
  });

  // Recherche instantanée
  document.getElementById('search-input').addEventListener('input', evt => {
    AppState.searchQuery = evt.target.value;
    renderItemList();
  });

  // Import manuel d'un registre (état vide + bouton "Importer un autre registre")
  const fileInput = document.getElementById('file-input');
  document.querySelectorAll('.btn-import').forEach(btn => {
    btn.addEventListener('click', () => fileInput.click());
  });
  fileInput.addEventListener('change', evt => {
    const file = evt.target.files[0];
    if (file) handleFileInput(file);
    fileInput.value = ''; // permet de resélectionner le même fichier plus tard
  });

  // Glisser-déposer sur la zone d'état vide
  const dropzone = document.getElementById('dropzone');
  ['dragenter', 'dragover'].forEach(evtName => {
    dropzone.addEventListener(evtName, e => {
      e.preventDefault();
      dropzone.classList.add('drag-active');
    });
  });
  ['dragleave', 'drop'].forEach(evtName => {
    dropzone.addEventListener(evtName, e => {
      e.preventDefault();
      dropzone.classList.remove('drag-active');
    });
  });
  dropzone.addEventListener('drop', e => {
    const file = e.dataTransfer.files[0];
    if (file) handleFileInput(file);
  });
}

// =============================================================================
// DÉMARRAGE
// =============================================================================

document.addEventListener('DOMContentLoaded', init);
