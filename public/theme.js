// Applies the saved or system theme before the first paint (an external file, so the CSP needs no inline script).
(function () {
  let dark;
  try {
    const v = localStorage.getItem("javascript-excel-playground:theme");
    dark = v === "dark" || (v !== "light" && matchMedia("(prefers-color-scheme: dark)").matches);
  } catch {
    dark = false;
  }
  if (dark) document.documentElement.classList.add("dark");
})();
