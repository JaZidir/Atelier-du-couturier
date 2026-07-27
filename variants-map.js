<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Atelier du Couturier</title>
  <link rel="stylesheet" href="css/style.css">
</head>
<body>

  <div class="app">

    <!-- ================= HEADER ================= -->
    <header class="app-header">
      <div class="brand">
        <!-- Marque : aiguille et fil, signature discrète de l'atelier -->
        <svg class="brand-mark" viewBox="0 0 32 32" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M9 23 C9 23 20 21 24 13 C25.5 10 24.5 7 22 6.5 C19.5 6 17.5 8 18 10.5"
                stroke="currentColor" stroke-width="1.6" stroke-linecap="round" fill="none"/>
          <circle cx="22.5" cy="8.5" r="2.6" stroke="currentColor" stroke-width="1.6" fill="none"/>
          <path d="M6 26 C6.5 23.5 8 22 9 21.5" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/>
        </svg>
        <div class="brand-titles">
          <span class="brand-title">Atelier du Couturier</span>
          <span class="brand-subtitle">Registre de fabrication</span>
        </div>
      </div>

      <nav class="tabs">
        <button class="tab-button active" data-tab="catalogue">Catalogue</button>
        <button class="tab-button" data-tab="rachat">Rachat des matériaux</button>
      </nav>

      <div class="header-spacer"></div>

      <button class="btn-import">Importer un registre</button>
    </header>

    <!-- ================= ÉTAT VIDE (aucun registre chargé) ================= -->
    <div id="empty-state">
      <div class="dropzone" id="dropzone">
        <svg class="dropzone-icon" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M12 4v11m0 0-4-4m4 4 4-4" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/>
          <path d="M4 16v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
        <h2>Aucun registre chargé</h2>
        <p>
          Déposez ici votre fichier Excel (feuilles <strong>VENTES</strong> et <strong>MATERIAUX</strong>),
          ou placez-le dans <code>data/registre.xlsx</code> pour un chargement automatique
          au prochain démarrage.
        </p>
        <button class="btn-primary btn-import">Parcourir mon ordinateur</button>
      </div>
    </div>

    <!-- Entrée fichier cachée, réutilisée par les deux boutons "Importer" -->
    <input type="file" id="file-input" accept=".xlsx" class="hidden">

    <!-- ================= CONTENU PRINCIPAL ================= -->
    <main id="app-content" class="hidden">

      <!-- ---- Onglet Catalogue ---- -->
      <section id="panel-catalogue" class="tab-panel">
        <div class="catalogue-layout">

          <aside class="catalogue-sidebar">
            <div class="sidebar-search">
              <input type="text" id="search-input" class="search-input" placeholder="Rechercher un objet…">
            </div>
            <div id="category-list"></div>
            <div id="item-list"></div>
          </aside>

          <div id="detail-panel"></div>

        </div>
      </section>

      <!-- ---- Onglet Rachat des matériaux ---- -->
      <section id="panel-rachat" class="tab-panel hidden">
        <div class="rachat-layout">
          <div class="rachat-header">
            <h2>Rachat des matériaux</h2>
            <p>
              Modifiez le prix de rachat d'un matériau : le coût de fabrication, le bénéfice
              et la rentabilité de tous les objets du catalogue qui l'utilisent sont
              recalculés instantanément.
            </p>
          </div>
          <div class="rachat-table-wrap">
            <table class="rachat-table">
              <thead>
                <tr>
                  <th>Matériau</th>
                  <th>Prix de rachat</th>
                  <th>Quantité</th>
                  <th>Prix unitaire</th>
                  <th></th>
                </tr>
              </thead>
              <tbody id="rachat-tbody"></tbody>
            </table>
          </div>
        </div>
      </section>

    </main>

  </div>

  <!-- Notification discrète -->
  <div id="toast" class="toast"></div>

  <!-- ================= SCRIPTS ================= -->
  <script src="lib/xlsx.full.min.js"></script>
  <script src="js/excel.js"></script>
  <script src="js/images-map.js"></script>
  <script src="js/variants-map.js"></script>
  <script src="js/app.js"></script>
</body>
</html>
