/** Instant tactile + visual helpers for mobile taps. */

export const pressable =
  "touch-manipulation select-none transition-[transform,filter,background-color] duration-75 ease-out active:scale-[0.96] active:brightness-75";

export function tapFeedback(): void {
  try {
    navigator.vibrate?.(12);
  } catch {
    /* ignore */
  }
}
