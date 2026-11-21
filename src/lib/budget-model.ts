export type LocationType = "standalone" | "mall" | null;
export type RoofStatus = "new" | "old" | null;
export type MaintenanceResponsibility = "operator" | "property" | "shared" | null;

export type CinemaFactorKey =
  | "attendance"
  | "areaSqm"
  | "age"
  | "renovation"
  | "locationType"
  | "roofStatus"
  | "maintenanceResponsibility"
  | string;

export interface Cinema {
  id: string;
  name: string;
  attendance: number | null;
  areaSqm: number | null;
  openingYear: number | null;
  renovated: boolean | null;
  locationType: LocationType;
  roofStatus: RoofStatus;
  maintenanceResponsibility: MaintenanceResponsibility;
  factorScores: Record<string, number>;
  note?: string;
}

export interface PriorityFactor {
  id: CinemaFactorKey;
  label: string;
  weight: number;
  enabled: boolean;
  kind: "number" | "category" | "manual";
  direction?: "higher" | "lower";
}

export interface HistoricalCategory {
  id: string;
  name: string;
  previousYearSpend: number;
}

export interface CategoryAllocation {
  categoryId: string;
  categoryName: string;
  previousYearSpend: number;
  share: number;
  amountCents: number;
}

export interface CinemaAllocation {
  cinemaId: string;
  coefficient: number;
  share: number;
  amountCents: number;
  categories: CategoryAllocation[];
}

export const DEFAULT_CINEMA_NAMES = [
  "TSC Beinasco (TO) TOF0001",
  "TSC Belpasso (CT) CAT0001",
  "TSC Bologna (BO) BOL0001",
  "TSC Casamassima (BA) BAR0001",
  "TSC Catanzaro (CZ) CTZ0001",
  "TSC Cerro Maggiore (MI) CER0001",
  "TSC Corciano (PG) PGR0001",
  "TSC Firenze (FI) NOV0001",
  "TSC Genova (GE) GEN0001",
  "TSC Grosseto (GR) GRO0001",
  "TSC Guidonia (RM) GUI0001",
  "TSC Lamezia Terme (CZ) LMZ0001",
  "TSC Limena (PD) PAD0001",
  "TSC Livorno (LI) LIV0001",
  "TSC Lugagnano (VR) VER0001",
  "TSC Montebello (PV) MON0001",
  "TSC Montesilvano (PE) PES0001",
  "TSC Napoli (NA) NAP0001",
  "TSC Nola (NA) NOL0001",
  "TSC Parma Barilla (PR) PAR0001",
  "TSC Parma Campus (PR) PAC0001",
  "TSC Pradamano (UD) UDI0001",
  "TSC Quartucciu (CA) CAG0001",
  "TSC Roma Moderno (RM) RMM0001",
  "TSC Roma Parco de' Medici (RM) RMP0001",
  "TSC Rozzano (MI) ROZ0001",
  "TSC Salerno (SA) SAL0001",
  "TSC Sestu (CA) SES0001",
  "TSC Silea (TV) TRE0001",
  "TSC Surbo (LE) SUR0001",
  "TSC Terni (TR) TER0001",
  "TSC Torino (TO) TOR0001",
  "TSC Torri di Quartesolo (VI) VIC0001",
  "TSC Trieste (TS) TRI0001",
  "TSC Vimercate (MB) VIM0001",
  "TSC Odeon Milano",
];

export function createDefaultCinemas(): Cinema[] {
  return DEFAULT_CINEMA_NAMES.map((name, index) => ({
    id: `cinema-${index + 1}`,
    name,
    attendance: null,
    areaSqm: null,
    openingYear: null,
    renovated: null,
    locationType: null,
    roofStatus: null,
    maintenanceResponsibility: null,
    factorScores: {},
  }));
}

export const DEFAULT_PRIORITY_FACTORS: PriorityFactor[] = [
  { id: "attendance", label: "Presenze annuali", weight: 1, enabled: true, kind: "number", direction: "higher" },
  { id: "areaSqm", label: "Superficie (m²)", weight: 1, enabled: true, kind: "number", direction: "higher" },
  { id: "age", label: "Anzianità del cinema", weight: 1, enabled: true, kind: "number", direction: "higher" },
  { id: "renovation", label: "Ristrutturazione", weight: 1, enabled: true, kind: "category" },
  { id: "locationType", label: "Standalone / centro commerciale", weight: 1, enabled: true, kind: "category" },
  { id: "roofStatus", label: "Roof", weight: 1, enabled: true, kind: "category" },
  { id: "maintenanceResponsibility", label: "Responsabilità manutenzione", weight: 1, enabled: true, kind: "category" },
];

const MIN_SCORE = 1;
const MAX_SCORE = 5;
const NEUTRAL_SCORE = 3;

function numericValue(value: number | null | undefined): number | null {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

function normalizedScores(
  values: Array<number | null>,
  direction: "higher" | "lower",
): number[] {
  const present = values.filter((value): value is number => value !== null);
  if (present.length === 0) return values.map(() => NEUTRAL_SCORE);

  const min = Math.min(...present);
  const max = Math.max(...present);
  if (min === max) return values.map((value) => value === null ? NEUTRAL_SCORE : NEUTRAL_SCORE);

  return values.map((value) => {
    if (value === null) return NEUTRAL_SCORE;
    const fraction = direction === "higher"
      ? (value - min) / (max - min)
      : (max - value) / (max - min);
    return MIN_SCORE + fraction * (MAX_SCORE - MIN_SCORE);
  });
}

function automaticScore(cinema: Cinema, factor: PriorityFactor): number {
  switch (factor.id) {
    case "renovation":
      return cinema.renovated === null ? NEUTRAL_SCORE : cinema.renovated ? 1 : 5;
    case "roofStatus":
      return cinema.roofStatus === null ? NEUTRAL_SCORE : cinema.roofStatus === "old" ? 5 : 1;
    case "maintenanceResponsibility":
      if (cinema.maintenanceResponsibility === null) return NEUTRAL_SCORE;
      if (cinema.maintenanceResponsibility === "operator") return 5;
      if (cinema.maintenanceResponsibility === "property") return 1;
      return NEUTRAL_SCORE;
    case "locationType":
      return NEUTRAL_SCORE;
    default:
      return NEUTRAL_SCORE;
  }
}

function factorScores(cinemas: Cinema[], factors: PriorityFactor[]): Map<string, number[]> {
  const scores = new Map<string, number[]>();
  for (const factor of factors) {
    let values: Array<number | null>;
    if (factor.id === "attendance") {
      values = normalizedScores(cinemas.map((cinema) => numericValue(cinema.attendance)), factor.direction ?? "higher");
    } else if (factor.id === "areaSqm") {
      values = normalizedScores(cinemas.map((cinema) => numericValue(cinema.areaSqm)), factor.direction ?? "higher");
    } else if (factor.id === "age") {
      const currentYear = new Date().getFullYear();
      values = normalizedScores(
        cinemas.map((cinema) => {
          const openingYear = numericValue(cinema.openingYear);
          return openingYear === null ? null : currentYear - openingYear;
        }),
        factor.direction ?? "higher",
      );
    } else {
      values = cinemas.map((cinema) => {
        const manual = cinema.factorScores[factor.id];
        if (typeof manual === "number" && Number.isFinite(manual)) {
          return Math.min(MAX_SCORE, Math.max(MIN_SCORE, manual));
        }
        return automaticScore(cinema, factor);
      });
    }

    scores.set(factor.id, values.map((score, index) => {
      const manual = cinemas[index]?.factorScores[factor.id];
      if (typeof manual === "number" && Number.isFinite(manual)) {
        return Math.min(MAX_SCORE, Math.max(MIN_SCORE, manual));
      }
      return score ?? NEUTRAL_SCORE;
    }));
  }
  return scores;
}

/**
 * Splits an integer number of cents proportionally and assigns leftover cents
 * by largest remainder so every output adds up exactly to the input total.
 */
export function splitCents(totalCents: number, weights: number[]): number[] {
  if (!Number.isFinite(totalCents) || totalCents <= 0 || weights.length === 0) {
    return weights.map(() => 0);
  }

  const validWeights = weights.map((weight) => Math.max(0, Number.isFinite(weight) ? weight : 0));
  const weightTotal = validWeights.reduce((sum, weight) => sum + weight, 0);
  if (weightTotal <= 0) return weights.map(() => 0);

  const exact = validWeights.map((weight) => totalCents * weight / weightTotal);
  const result = exact.map(Math.floor);
  let remaining = totalCents - result.reduce((sum, cents) => sum + cents, 0);
  const order = exact
    .map((value, index) => ({ index, remainder: value - Math.floor(value) }))
    .sort((a, b) => b.remainder - a.remainder || a.index - b.index);

  for (let i = 0; i < remaining; i += 1) {
    result[order[i % order.length].index] += 1;
  }
  remaining = 0;
  return result;
}

export function calculateAllocations(
  cinemas: Cinema[],
  factors: PriorityFactor[],
  totalBudget: number,
  categories: HistoricalCategory[],
): CinemaAllocation[] {
  const usableFactors = factors.filter((factor) => factor.enabled && factor.weight > 0);
  const scores = factorScores(cinemas, usableFactors);
  const totalWeight = usableFactors.reduce((sum, factor) => sum + factor.weight, 0);
  const coefficients = cinemas.map((cinema, cinemaIndex) => {
    if (totalWeight === 0) return 0;
    const weightedScore = usableFactors.reduce((sum, factor) => {
      return sum + (scores.get(factor.id)?.[cinemaIndex] ?? NEUTRAL_SCORE) * factor.weight;
    }, 0);
    return weightedScore / totalWeight;
  });

  const budgetCents = Math.max(0, Math.round((Number.isFinite(totalBudget) ? totalBudget : 0) * 100));
  const cinemaCents = splitCents(budgetCents, coefficients);
  const categoryWeights = categories.map((category) => Math.max(0, category.previousYearSpend));

  return cinemas.map((cinema, index) => {
    const coefficient = coefficients[index] ?? 0;
    const share = coefficients.reduce((sum, value) => sum + value, 0) > 0
      ? coefficient / coefficients.reduce((sum, value) => sum + value, 0)
      : 0;
    const amounts = splitCents(cinemaCents[index] ?? 0, categoryWeights);
    const previousYearTotal = categoryWeights.reduce((sum, amount) => sum + amount, 0);
    return {
      cinemaId: cinema.id,
      coefficient,
      share,
      amountCents: cinemaCents[index] ?? 0,
      categories: categories.map((category, categoryIndex) => ({
        categoryId: category.id,
        categoryName: category.name,
        previousYearSpend: category.previousYearSpend,
        share: previousYearTotal > 0 ? categoryWeights[categoryIndex] / previousYearTotal : 0,
        amountCents: amounts[categoryIndex] ?? 0,
      })),
    };
  });
}

export function centsToEuros(cents: number): number {
  return cents / 100;
}
