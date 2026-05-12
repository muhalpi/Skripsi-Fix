const POINTS_PER_CM = 28.3464567;

export function cmToPoints(cm: number): number {
  return cm * POINTS_PER_CM;
}

export function pointsToCm(points: number): number {
  return points / POINTS_PER_CM;
}

export function almostEqual(a: number, b: number, tolerance = 0.25): boolean {
  return Math.abs(a - b) <= tolerance;
}
