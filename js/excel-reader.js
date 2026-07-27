/**
 * excel-reader.js
 * Lecteur de fichiers Excel (.xlsx) utilisant SheetJS
 * Extrait et normalise les données du registre
 */

const ExcelReader = {
  utils: {
    normalize(str) {
      return (str ?? '').toLowerCase().trim().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    },
    
    normalizeHeader(str) {
      return this.normalize(str);
    }
  },

  loadFromDefaultPath: async function(path) {
    try {
      const response = await fetch(path);
      if (!response.ok) throw new Error('Fichier non trouvé');
      const buffer = await response.arrayBuffer();
      return this.loadFromArrayBuffer(buffer);
    } catch (err) {
      throw new Error(`Impossible de charger ${path}: ${err.message}`);
    }
  },

  loadFromArrayBuffer: function(buffer) {
    if (typeof XLSX === 'undefined') {
      throw new Error('SheetJS (XLSX) n\'est pas chargé');
    }

    try {
      const workbook = XLSX.read(buffer, { type: 'array' });
      const materiaux = this.readMateriauxSheet(workbook);
      const objets = this.readVentesSheet(workbook, materiaux);
      
      return { materiaux, objets };
    } catch (err) {
      throw new Error(`Erreur lors de la lecture: ${err.message}`);
    }
  },

  readMateriauxSheet: function(workbook) {
    const sheetName = workbook.SheetNames.find(n => 
      this.utils.normalizeHeader(n).includes('materia')
    );
    
    if (!sheetName) {
      console.warn('Feuille MATERIAUX non trouvée');
      return [];
    }

    const ws = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

    const materiaux = [];
    const seen = new Set();

    for (const row of rows) {
      // Recherche le nom du matériau dans les colonnes
      let name = '';
      let buyPrice = 0;
      let quantity = 1;

      for (const [key, value] of Object.entries(row)) {
        const normalizedKey = this.utils.normalizeHeader(key);
        
        if (normalizedKey.includes('objet') || normalizedKey.includes('materia') || normalizedKey.includes('nom')) {
          if (value && typeof value === 'string' && value.trim()) {
            name = value.trim();
          }
        }
        
        if (normalizedKey.includes('prix') && normalizedKey.includes('achat')) {
          buyPrice = parseFloat(value) || 0;
        }
        
        if (normalizedKey.includes('quantite') || normalizedKey.includes('quantité')) {
          quantity = parseFloat(value) || 1;
        }
      }

      if (!name || seen.has(name.toLowerCase())) continue;
      seen.add(name.toLowerCase());

      const unitPrice = quantity > 0 ? buyPrice / quantity : buyPrice;

      materiaux.push({
        id: name.toLowerCase().replace(/[^a-z0-9]+/g, '-'),
        name,
        buyPrice,
        quantity,
        unitPrice,
        originalBuyPrice: buyPrice,
        customBuyPrice: null
      });
    }

    console.log(`Chargé ${materiaux.length} matériaux`);
    return materiaux;
  },

  readVentesSheet: function(workbook, materiaux) {
    const sheetName = workbook.SheetNames.find(n => 
      this.utils.normalizeHeader(n).includes('vente')
    );
    
    if (!sheetName) {
      console.warn('Feuille VENTES non trouvée');
      return [];
    }

    const ws = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

    const objets = [];
    const categoryKeywords = {
      'Accessoires': ['bracelet', 'bague', 'collar', 'amulet', 'charm', 'ring', 'necklace', 'bijou'],
      'Robes': ['robe', 'dress', 'gown'],
      'Fine Clothes': ['fine', 'arcana', 'aureate', 'bardic', 'blood', 'blue', 'burnished', 'charcoal'],
      'Capuches': ['capuche', 'cowl', 'hood'],
      'Vêtements': ['vetement', 'clothes', 'garment', 'outfit', 'tunic', 'tunicque'],
      'Bottes': ['bottes', 'boots'],
      'Chaussures': ['chaussure', 'shoes']
    };

    for (const row of rows) {
      let name = '';
      let sellPrice = 0;
      let category = 'Autres';
      let recipe = [];
      let excelCost = 0;

      for (const [key, value] of Object.entries(row)) {
        const normalizedKey = this.utils.normalizeHeader(key);

        if (normalizedKey.includes('objet') || normalizedKey.includes('nom')) {
          if (value && typeof value === 'string' && value.trim()) {
            name = value.trim();
          }
        }

        if (normalizedKey.includes('prix') && normalizedKey.includes('vente')) {
          sellPrice = parseFloat(value) || 0;
        }

        if (normalizedKey.includes('cout') || normalizedKey.includes('coût')) {
          excelCost = parseFloat(value) || 0;
        }
      }

      if (!name) continue;

      // Déterminer la catégorie
      const nameLower = this.utils.normalize(name);
      for (const [cat, keywords] of Object.entries(categoryKeywords)) {
        if (keywords.some(kw => nameLower.includes(kw))) {
          category = cat;
          break;
        }
      }

      // Extraire la recette si disponible
      for (const [key, value] of Object.entries(row)) {
        const normalizedKey = this.utils.normalizeHeader(key);
        if (normalizedKey.match(/materia[ux]?\s*\d+/) || normalizedKey.includes('materiel')) {
          if (value) {
            const materialName = String(value).trim();
            const qty = 1; // Quantité par défaut
            const material = materiaux.find(m => m.name.toLowerCase() === materialName.toLowerCase());
            if (material) {
              recipe.push({
                materialId: material.id,
                materialName: materialName,
                quantity: qty
              });
            }
          }
        }
      }

      objets.push({
        id: name.toLowerCase().replace(/[^a-z0-9]+/g, '-'),
        name,
        category,
        sellPrice,
        recipe,
        excelCost,
        imageKey: name.toLowerCase().replace(/[^a-z0-9]+/g, '-')
      });
    }

    console.log(`Chargé ${objets.length} objets`);
    return objets;
  }
};

// Vérifier que le module est bien défini
if (typeof ExcelReader === 'undefined') {
  throw new Error('ExcelReader non défini');
}
