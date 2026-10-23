// Colle ici l'URL de l'application Web Google Apps Script (voir README).
// Laisse vide pour un affichage en lecture seule à partir de data.js.
window.STONETROOPERS_CONFIG = {
  API_URL: "https://script.google.com/macros/s/AKfycbxQ5z4Yp7GfO6ynjU8dhStBHmZ8AaHLGJpjUdQ58v_ElvAOQEf5iTAWJc1RLWBpAncpMQ/exec",
  // Sections de la page : true = affichée, false = masquée
  AFFICHER_VACANCES: false,
  AFFICHER_EQUIPE: false,
  AFFICHER_STATS: true,        // visible seulement pour les EDITEURS
  // Joueurs qui voient les boutons « + Match », ✏️ et la saisie de composition quand ils sont choisis dans « Je suis »
  EDITEURS: ["Pascal"],
  // Export calendrier (.ics) : durée d'un match en minutes, et lieu (facultatif, ex. "Patinoire de …")
  DUREE_MATCH_MIN: 120,
  LIEU: "",
  // Composition proposée par défaut sur un match terminé (ordre 1 → 4) et skip habituel
  COMPO_HABITUELLE: ["Suettania", "Pascal", "Océane", "Adélie"],
  SKIP_HABITUEL: "Suettania",
  // Formation de remplacement quand un joueur est absent ("" = place libre pour un autre joueur présent)
  COMPO_SI_ABSENT: {
    "Suettania": { ordre: ["Pascal", "", "Océane", "Adélie"], skip: "Océane" }
  }
};
