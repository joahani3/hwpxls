export const MAPPING_COLORS = [
  { bg: "#dbeafe", border: "#93c5fd", badge: "#2563eb", text: "#1e40af" },
  { bg: "#d1fae5", border: "#6ee7b7", badge: "#059669", text: "#065f46" },
  { bg: "#fde68a", border: "#fbbf24", badge: "#d97706", text: "#78350f" },
  { bg: "#fce7f3", border: "#f9a8d4", badge: "#db2777", text: "#831843" },
  { bg: "#ede9fe", border: "#c4b5fd", badge: "#7c3aed", text: "#4c1d95" },
  { bg: "#ffedd5", border: "#fdba74", badge: "#ea580c", text: "#7c2d12" },
  { bg: "#cffafe", border: "#67e8f9", badge: "#0891b2", text: "#164e63" },
  { bg: "#fef9c3", border: "#fde047", badge: "#ca8a04", text: "#713f12" },
];

export function mappingColor(idx: number) {
  return MAPPING_COLORS[idx % MAPPING_COLORS.length];
}
