/**
 * excel.js
 * ---------------------------------------------------------------------------
 * Module responsable de la lecture du registre Excel (.xlsx) et de sa
 * transformation en structures de données exploitables par l'application.
 *
 * Ce module ne contient AUCUNE donnée en dur : tout provient du fichier
 * Excel fourni par l'utilisateur (feuilles "VENTES" et "MATERIAUX").
 *
 * Il est volontairement tolérant sur la mise en forme du registre :
 * - la ligne d'en-tête n'est pas forcément la première ligne de la feuille
 * - les intitulés de colonnes peuvent varier légèrement (accents, casse,
 *   espaces) : on les normalise avant de les comparer
 * - le nombre de couples (matériau / quantité) peut varier de 1 à 6
 *
 * Dépendance : lib/xlsx.full.min.js (SheetJS), chargé avant ce fichier.
 * ---------------------------------------------------------------------------
 */

const ExcelReader = (() => {

  // ---------------------------------------------------------------------
  // Constantes
  // ---------------------------------------------------------------------

  const SHEET_NAMES = {
    VENTES: 'VENTES',
    MATERIAUX: 'MATERIAUX',
    ACCUEIL: 'ACCUEIL' // volontairement ignorée
  };

  // Dictionnaire d'inférence de catégorie utilisé uniquement si le registre
  // ne contient pas lui-même de colonne "Catégorie". L'ORDRE compte : la
  // première catégorie dont un mot-clé correspond l'emporte (ex: "Robe à
  // Capuchon" doit rester une Robe, pas atterrir dans Capuches).
  // Modifiable librement : il suffit d'ajouter des entrées pour affiner
  // la classification au fur et à mesure que de nouveaux objets arrivent.
  const CATEGORY_KEYWORDS = [
    { category: 'Bottes',       keywords: ['botte'] },
    { category: 'Chaussures',   keywords: ['chaussure', 'sandal', 'soulier'] },
    { category: 'Fine Clothes', keywords: ['fine clothes'] },
    { category: 'Robes',        keywords: ['robe', 'tunic', 'tunique', 'surcoat'] },
    { category: 'Capuches',     keywords: ['capuche', 'capuchon', 'cowl', 'chapeau', 'bonnet', 'toque'] },
    { category: 'Vêtements',    keywords: ['vetement', 'clothes', 'habit', 'pantalon'] },
    { category: 'Accessoires',  keywords: [
        'scarf', 'echarpe', 'foulard', 'gaiter', 'satchel', 'sac', 'pouch',
        'mantle', 'manteau', 'cape', 'cloak', 'mask', 'masque', 'backpack',
        'resource', 'bandage', 'ceinture', 'tablier'
      ] },
  ];
  const DEFAULT_CATEGORY = 'Autres';

  // ---------------------------------------------------------------------
  // Utilitaires de normalisation de texte
  // ---------------------------------------------------------------------

  /** Retire les accents, met en minuscule et supprime les espaces superflus. */
  function normalize(str) {
    if (str === null || str === undefined) return '';
    return str
      .toString()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLowerCase()
      .trim();
  }

  /** Convertit une valeur de cellule en nombre exploitable (gère "12,5", "12%", espaces, cellules vides). */
  function toNumber(value) {
    if (value === null || value === undefined || value === '') return 0;
    if (typeof value === 'number') return value;
    let str = value.toString().trim();
    const isPercent = str.includes('%');
    str = str.replace('%', '').replace(/\s/g, '').replace(',', '.');
    const num = parseFloat(str);
    if (isNaN(num)) return 0;
    return isPercent ? num / 100 : num;
  }

  /** Transforme un nom d'objet/matériau en identifiant technique stable (slug). */
  function slugify(name) {
    return normalize(name).replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '');
  }

  // ---------------------------------------------------------------------
  // Détection de la ligne d'en-tête
  // ---------------------------------------------------------------------

  /**
   * Cherche, dans les 15 premières lignes d'une feuille, celle qui contient
   * une cellule correspondant à l'un des mots-clés attendus (ex: "objet" ou
   * "nom", selon la convention de nommage du registre).
   * Retourne l'index de ligne (0-based) ou 0 si rien n'est trouvé.
   */
  function findHeaderRowIndex(sheet, requiredKeywords) {
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
    const maxScan = Math.min(rows.length, 15);
    for (let i = 0; i < maxScan; i++) {
      const row = rows[i] || [];
      const hasKeyword = row.some(cell => requiredKeywords.includes(normalize(cell)));
      if (hasKeyword) return i;
    }
    return 0; // fallback : on suppose que la première ligne est l'en-tête
  }

  /** Lit une feuille sous forme de tableau d'objets {enTête: valeur}, à partir de la vraie ligne d'en-tête. */
  function readSheetAsObjects(sheet, requiredKeywords) {
    const headerRow = findHeaderRowIndex(sheet, requiredKeywords);
    return XLSX.utils.sheet_to_json(sheet, { range: headerRow, defval: '' });
  }

  // ---------------------------------------------------------------------
  // Détection des colonnes par mots-clés (insensible aux accents/casse)
  // ---------------------------------------------------------------------

  /** Retourne, pour un objet-ligne donné, la clé de colonne correspondant au premier mot-clé trouvé. */
  function findColumnKey(rowKeys, predicate) {
    return rowKeys.find(predicate) || null;
  }

  // ---------------------------------------------------------------------
  // Parsing de la feuille MATERIAUX
  // ---------------------------------------------------------------------

  function parseMateriaux(workbook) {
    const sheetName = Object.keys(workbook.Sheets).find(
      n => normalize(n) === normalize(SHEET_NAMES.MATERIAUX)
    );
    if (!sheetName) {
      throw new Error(`Feuille "${SHEET_NAMES.MATERIAUX}" introuvable dans le registre.`);
    }

    const sheet = workbook.Sheets[sheetName];
    const rows = readSheetAsObjects(sheet, ['objet', 'nom']);

    return rows
      .filter(row => Object.values(row).some(v => v !== '')) // ignore lignes vides
      .map(row => {
        const keys = Object.keys(row);
        const nameKey = findColumnKey(keys, k => {
          const n = normalize(k);
          return n.includes('objet') || n.includes('nom');
        });
        const buyPriceKey = findColumnKey(keys, k => normalize(k).includes('prix') && normalize(k).includes('achat'));
        const qtyKey = findColumnKey(keys, k => {
          const n = normalize(k);
          return (n.includes('quantite') || n === 'qte') && !n.includes('prix');
        });
        const unitPriceKey = findColumnKey(keys, k => normalize(k).includes('prix') && normalize(k).includes('unit'));

        const name = nameKey ? row[nameKey].toString().trim() : '';
        const buyPrice = buyPriceKey ? toNumber(row[buyPriceKey]) : 0;
        const quantity = qtyKey ? toNumber(row[qtyKey]) : 1;
        // Si un prix unitaire est déjà calculé dans le registre, on le garde comme référence,
        // sinon on le déduit de prix d'achat / quantité.
        const declaredUnitPrice = unitPriceKey ? toNumber(row[unitPriceKey]) : 0;
        const computedUnitPrice = quantity > 0 ? buyPrice / quantity : buyPrice;

        if (!name) return null;

        return {
          id: slugify(name),
          name,
          buyPrice,
          quantity: quantity || 1,
          unitPrice: declaredUnitPrice > 0 ? declaredUnitPrice : computedUnitPrice,
          // Prix de rachat modifiable par l'utilisateur dans l'onglet dédié.
          // Initialisé à la valeur du registre ; conservé séparément pour
          // permettre un "réinitialiser" fidèle à la source Excel.
          originalBuyPrice: buyPrice,
          customBuyPrice: null // null = pas de surcharge, on utilise buyPrice
        };
      })
      .filter(Boolean);
  }

  // ---------------------------------------------------------------------
  // Parsing de la feuille VENTES
  // ---------------------------------------------------------------------

  function parseVentes(workbook) {
    const sheetName = Object.keys(workbook.Sheets).find(
      n => normalize(n) === normalize(SHEET_NAMES.VENTES)
    );
    if (!sheetName) {
      throw new Error(`Feuille "${SHEET_NAMES.VENTES}" introuvable dans le registre.`);
    }

    const sheet = workbook.Sheets[sheetName];
    const rows = readSheetAsObjects(sheet, ['objet', 'nom']);

    // --- Étape 1 : extraction brute de chaque ligne -----------------------
    const rawItems = rows
      .filter(row => Object.values(row).some(v => v !== ''))
      .map(row => {
        const keys = Object.keys(row);

        // "Objet" est le nom usuel dans ce registre ; "Nom" est accepté par
        // souplesse pour d'autres registres. On exclut les colonnes de
        // recette ("Materiel N") qui pourraient contenir "nom" par hasard.
        const nameKey = findColumnKey(keys, k => {
          const n = normalize(k);
          return (n.includes('objet') || n.includes('nom'))
            && !n.includes('materiau') && !n.includes('matiere') && !n.includes('materiel');
        });
        const sellPriceKey = findColumnKey(keys, k => normalize(k).includes('prix') && normalize(k).includes('vente'));
        const costKey = findColumnKey(keys, k => normalize(k).includes('cout'));
        const marginKey = findColumnKey(keys, k => normalize(k).includes('marge'));
        // "% de marge" contient aussi "marge" : on la distingue par la présence du signe %.
        const percentKey = findColumnKey(keys, k => {
          const n = normalize(k);
          return n.includes('%') || n.includes('rentab') || n.includes('pourcent');
        });
        const categoryKey = findColumnKey(keys, k => normalize(k).includes('categorie'));

        // Colonnes de matériaux ("Materiel N") et de quantités ("Quantitée N") :
        // récupérées dans l'ordre d'apparition, puis associées deux à deux.
        const materialKeys = keys.filter(k => {
          const n = normalize(k);
          return (n.includes('materiau') || n.includes('matiere') || n.includes('materiel'))
            && !n.includes('nom') && !n.includes('objet');
        });
        const quantityKeys = keys.filter(k => {
          const n = normalize(k);
          return (n.includes('quantite') || n.includes('qte')) && !n.includes('prix');
        });

        const name = nameKey ? row[nameKey].toString().trim() : '';
        if (!name) return null;

        const recipe = [];
        for (const materialKey of materialKeys) {
          const materialName = row[materialKey] ? row[materialKey].toString().trim() : '';
          if (!materialName) continue;

          // trouver l'index de la colonne matériau dans l'ordre des clés
          const materialIdx = keys.indexOf(materialKey);
          let qty = 0;

          // chercher la colonne quantité la plus proche à droite
          for (let j = materialIdx + 1; j < keys.length; j++) {
            const candidate = keys[j];
            if (quantityKeys.includes(candidate)) {
              qty = toNumber(row[candidate]);
              break;
            }
            // si on rencontre un autre "matériau" avant une quantité, on arrête
            if (materialKeys.includes(candidate)) break;
          }

          if (qty <= 0) continue; // ignore quantités invalides
          recipe.push({ materialId: slugify(materialName), materialName, quantity: qty });
        }

        const sellPrice = sellPriceKey ? toNumber(row[sellPriceKey]) : 0;
        const excelCost = costKey ? toNumber(row[costKey]) : 0;
        const excelMargin = marginKey ? toNumber(row[marginKey]) : 0;
        const excelPercent = percentKey ? toNumber(row[percentKey]) : 0;

        const category = categoryKey && row[categoryKey]
          ? row[categoryKey].toString().trim()
          : inferCategory(name);

        return { name, category, sellPrice, recipe, excelCost, excelMargin, excelPercent };
      })
      .filter(Boolean);

    // --- Étape 2 : identifiants uniques + libellé des recettes alternatives ---
    // Certains objets du registre partagent le même nom avec des recettes
    // différentes (ex: "Bottes" x3, un tarif/coût par recette). On les
    // conserve toutes, en distinguant clairement chaque variante plutôt que
    // de les faire silencieusement écraser les unes les autres.
    const occurrenceCount = new Map();
    rawItems.forEach(item => {
      const baseId = slugify(item.name);
      occurrenceCount.set(baseId, (occurrenceCount.get(baseId) || 0) + 1);
    });

    const occurrenceIndex = new Map();
    return rawItems.map(item => {
      const baseId = slugify(item.name);
      const total = occurrenceCount.get(baseId);
      const index = (occurrenceIndex.get(baseId) || 0) + 1;
      occurrenceIndex.set(baseId, index);

      const isVariant = total > 1;
      return {
        id: isVariant ? `${baseId}-${index}` : baseId,
        // Identifiant de base (sans suffixe de variante), utilisé pour
        // retrouver l'image : deux recettes du même objet partagent la
        // même illustration.
        imageKey: baseId,
        name: isVariant ? `${item.name} (recette ${index}/${total})` : item.name,
        category: item.category,
        sellPrice: item.sellPrice,
        recipe: item.recipe,
        excelCost: item.excelCost,
        excelMargin: item.excelMargin,
        excelPercent: item.excelPercent
      };
    });
  }

  /** Déduit une catégorie à partir du nom de l'objet, quand le registre n'en fournit pas. */
  function inferCategory(name) {
    const n = normalize(name);
    for (const entry of CATEGORY_KEYWORDS) {
      if (entry.keywords.some(kw => n.includes(normalize(kw)))) {
        return entry.category;
      }
    }
    return DEFAULT_CATEGORY;
  }

  // ---------------------------------------------------------------------
  // API publique
  // ---------------------------------------------------------------------

  /**
   * Charge un registre à partir d'un ArrayBuffer (issu d'un <input type="file">
   * ou d'un fetch() sur data/registre.xlsx) et retourne les données normalisées.
   */
  function loadFromArrayBuffer(arrayBuffer) {
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const materiaux = parseMateriaux(workbook);
    const objets = parseVentes(workbook);
    return { materiaux, objets };
  }

  /** Tente de charger automatiquement le registre embarqué dans data/registre.xlsx. */
  async function loadFromDefaultPath(path = 'data/registre.xlsx') {
    const response = await fetch(path);
    if (!response.ok) {
      throw new Error('Aucun registre trouvé à cet emplacement.');
    }
    const buffer = await response.arrayBuffer();
    return loadFromArrayBuffer(buffer);
  }

  return {
    loadFromArrayBuffer,
    loadFromDefaultPath,
    // Exposés pour tests / réutilisation (ex: futur module de stock).
    utils: { normalize, toNumber, slugify }
  };
})();
