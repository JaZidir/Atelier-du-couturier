/**
 * variants-map.js
 * ---------------------------------------------------------------------------
 * Déclinaisons visuelles (couleurs, styles...) d'un même objet, quand elles
 * ne sont PAS distinguées dans le registre Excel — par exemple une "Robe"
 * qui existe en bleu, gris, marron, etc. mais n'a qu'une seule ligne dans
 * VENTES.
 *
 * Si un objet a une entrée ici, l'application affiche un menu déroulant
 * dans le panneau de détail permettant de choisir la variante ; le nom
 * choisi (label) s'affiche dans ce menu, et les images "Porté homme" /
 * "Porté femme" changent en conséquence.
 *
 * Si un objet N'A PAS d'entrée ici, aucun menu déroulant n'apparaît : ses
 * images restent simplement images/<imageKey>-homme.png et
 * images/<imageKey>-femme.png comme avant.
 *
 * ---------------------------------------------------------------------------
 * COMMENT AJOUTER UNE VARIANTE (aucune connaissance en code nécessaire) :
 *
 * 1. Trouvez l'imageKey de l'objet concerné : c'est le nom de fichier que
 *    l'application cherche par défaut, visible dans la colonne C ou D de
 *    data/images_a_completer.xlsx (ex: pour "images/robe-bleu-homme.png",
 *    l'imageKey est "robe-bleu").
 *
 * 2. Ajoutez (ou complétez) une ligne dans VARIANTS_MAP ci-dessous, en
 *    copiant EXACTEMENT ce modèle (attention aux virgules et guillemets) :
 *
 *    "robe-bleu": [
 *      { label: "Bleu",   homme: "images/robe-bleu-homme.png",   femme: "images/robe-bleu-femme.png" },
 *      { label: "Grise",  homme: "images/robe-grise-homme.png",  femme: "images/robe-grise-femme.png" },
 *      { label: "Marron", homme: "images/robe-marron-homme.png", femme: "images/robe-marron-femme.png" }
 *    ],
 *
 * 3. Chaque variante doit avoir : un "label" (le nom qui s'affichera dans
 *    le menu déroulant), un "homme" et un "femme" (chemins des fichiers
 *    à déposer dans le dossier images/). Vous pouvez mettre autant de
 *    variantes que vous voulez, dans l'ordre que vous voulez — la
 *    première de la liste s'affiche par défaut.
 *
 * 4. Une virgule sépare chaque ligne "imageKey": [...] SAUF la toute
 *    dernière. Enregistrez le fichier et rechargez la page.
 * ---------------------------------------------------------------------------
 */

const VARIANTS_MAP = {

  // Exemple (à adapter ou supprimer) :
  // "robe-bleu": [
  //   { label: "Bleu",   homme: "images/robe-bleu-homme.png",   femme: "images/robe-bleu-femme.png" },
  //   { label: "Grise",  homme: "images/robe-grise-homme.png",  femme: "images/robe-grise-femme.png" },
  //   { label: "Marron", homme: "images/robe-marron-homme.png", femme: "images/robe-marron-femme.png" }
  // ],

};
