import { useCallback, useEffect, useState } from "react";

export type ThemeChoice = "light" | "dark" | "system";
export const THEME_KEY = "javascript-excel-playground:theme";

function stored(): ThemeChoice {
  try {
    const v = localStorage.getItem(THEME_KEY);
    return v === "light" || v === "dark" ? v : "system";
  } catch {
    return "system";
  }
}

const systemDark = (): boolean => typeof matchMedia === "function" && matchMedia("(prefers-color-scheme: dark)").matches;

/** Light, dark or the system setting; applied as the `dark` class on <html>. */
export function useTheme(): { choice: ThemeChoice; dark: boolean; setChoice: (c: ThemeChoice) => void; toggle: () => void } {
  const [choice, setChoiceState] = useState<ThemeChoice>(stored);
  const [sys, setSys] = useState(systemDark);

  useEffect(() => {
    if (typeof matchMedia !== "function") return;
    const mq = matchMedia("(prefers-color-scheme: dark)");
    const on = (): void => setSys(mq.matches);
    mq.addEventListener("change", on);
    return () => mq.removeEventListener("change", on);
  }, []);

  const dark = choice === "dark" || (choice === "system" && sys);
  useEffect(() => {
    document.documentElement.classList.toggle("dark", dark);
    document.documentElement.style.colorScheme = dark ? "dark" : "light";
  }, [dark]);

  const setChoice = useCallback((c: ThemeChoice) => {
    setChoiceState(c);
    try {
      if (c === "system") localStorage.removeItem(THEME_KEY);
      else localStorage.setItem(THEME_KEY, c);
    } catch {
      // ignore: the choice still applies to this visit
    }
  }, []);

  const toggle = useCallback(() => setChoice(dark ? "light" : "dark"), [dark, setChoice]);
  return { choice, dark, setChoice, toggle };
}
