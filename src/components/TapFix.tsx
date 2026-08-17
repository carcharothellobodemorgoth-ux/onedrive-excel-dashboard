"use client";

import { useEffect } from "react";

/** Lets iOS apply :active styles (needs a touchstart listener on the document). */
export function TapFix() {
  useEffect(() => {
    const noop = () => {};
    document.addEventListener("touchstart", noop, { passive: true });
    return () => document.removeEventListener("touchstart", noop);
  }, []);
  return null;
}
