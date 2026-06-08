// Configuration des 3 magasins CLOANE
// Utilisée pour le ticket, la signature des messages clients, etc.

export const MAGASINS = {
  SQUARE: {
    code: "SQUARE",
    nom: "CLOANE SQUARE",
    tel: "02 97 63 61 95",
    adresse: "14 rue Henri Navier, 56000 Vannes",
  },
  SIGNATURE: {
    code: "SIGNATURE",
    nom: "CLOANE SIGNATURE",
    tel: "02 97 41 87 20",
    adresse: "Rue Henri Navier, 56000 Vannes",
  },
  STORE: {
    code: "STORE",
    nom: "CLOANE STORE",
    tel: "02 97 63 16 64",
    adresse: "95 Av. de la Marne, 56000 Vannes (C.Cial Carrefour)",
  },
};

export const MAGASIN_CODES = ["SQUARE", "SIGNATURE", "STORE"];

export function magasinNom(code) {
  return MAGASINS[code]?.nom || code || "";
}
export function magasinTel(code) {
  return MAGASINS[code]?.tel || "";
}
