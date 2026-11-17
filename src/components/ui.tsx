// Small form and layout primitives shared by the panels.
import { useId, type ReactNode } from "react";

export function Card({ title, children, actions, className = "" }: { title?: ReactNode; children: ReactNode; actions?: ReactNode; className?: string }) {
  return (
    <section className={`rounded-lg border border-slate-200 bg-white p-4 shadow-sm dark:border-slate-700 dark:bg-slate-900 ${className}`}>
      {(title || actions) && (
        <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
          {title && <h2 className="text-base font-semibold text-slate-900 dark:text-slate-100">{title}</h2>}
          {actions && <div className="flex flex-wrap gap-2">{actions}</div>}
        </div>
      )}
      {children}
    </section>
  );
}

export function Toggle({ label, checked, onChange, hint }: { label: string; checked: boolean; onChange: (v: boolean) => void; hint?: string }) {
  const id = useId();
  return (
    <div className="flex items-start gap-2 py-1">
      <input id={id} type="checkbox" className="mt-1 h-4 w-4 accent-blue-700" checked={checked} onChange={(e) => onChange(e.target.checked)} />
      <label htmlFor={id} className="text-sm text-slate-800 dark:text-slate-200">
        {label}
        {hint && <span className="block text-xs text-slate-600 dark:text-slate-400">{hint}</span>}
      </label>
    </div>
  );
}

export function Select<T extends string>({ label, value, options, onChange }: { label: string; value: T; options: ReadonlyArray<readonly [T, string]>; onChange: (v: T) => void }) {
  const id = useId();
  return (
    <div className="flex flex-col gap-1 py-1">
      <label htmlFor={id} className="text-sm text-slate-800 dark:text-slate-200">
        {label}
      </label>
      <select id={id} value={value} onChange={(e) => onChange(e.target.value as T)} className="field w-full min-w-0">
        {options.map(([v, l]) => (
          <option key={v} value={v}>
            {l}
          </option>
        ))}
      </select>
    </div>
  );
}

export function NumberField({ label, value, onChange, min = 0, max, hint }: { label: string; value: number; onChange: (v: number) => void; min?: number; max?: number; hint?: string }) {
  const id = useId();
  return (
    <div className="flex flex-col gap-1 py-1">
      <label htmlFor={id} className="text-sm text-slate-800 dark:text-slate-200">
        {label}
        {hint && <span className="block text-xs text-slate-600 dark:text-slate-400">{hint}</span>}
      </label>
      <input
        id={id}
        type="number"
        inputMode="numeric"
        min={min}
        max={max}
        value={value}
        onChange={(e) => {
          const n = Number(e.target.value);
          if (Number.isFinite(n)) onChange(Math.max(min, Math.min(max ?? n, Math.trunc(n))));
        }}
        className="field"
      />
    </div>
  );
}

export function TextField({ label, value, onChange, hint, error }: { label: string; value: string; onChange: (v: string) => void; hint?: string; error?: string | null }) {
  const id = useId();
  const noteId = `${id}-note`;
  return (
    <div className="flex flex-col gap-1 py-1">
      <label htmlFor={id} className="text-sm text-slate-800 dark:text-slate-200">
        {label}
      </label>
      <input id={id} type="text" value={value} onChange={(e) => onChange(e.target.value)} aria-invalid={error ? true : undefined} aria-describedby={hint || error ? noteId : undefined} className={`field ${error ? "border-red-600 dark:border-red-500" : ""}`} />
      {(hint || error) && (
        <p id={noteId} className={`text-xs ${error ? "text-red-700 dark:text-red-300" : "text-slate-600 dark:text-slate-400"}`}>
          {error ?? hint}
        </p>
      )}
    </div>
  );
}

export function Button({ children, onClick, variant = "secondary", disabled, title, type = "button", ariaKeyshortcuts }: { children: ReactNode; onClick?: () => void; variant?: "primary" | "secondary" | "ghost"; disabled?: boolean; title?: string; type?: "button" | "submit"; ariaKeyshortcuts?: string }) {
  const styles = {
    primary: "bg-blue-700 text-white hover:bg-blue-800 disabled:bg-slate-400",
    secondary: "border border-slate-300 bg-white text-slate-900 hover:bg-slate-100 dark:border-slate-600 dark:bg-slate-800 dark:text-slate-100 dark:hover:bg-slate-700",
    ghost: "text-slate-700 hover:bg-slate-100 dark:text-slate-200 dark:hover:bg-slate-800",
  }[variant];
  return (
    <button type={type} onClick={onClick} disabled={disabled} title={title} aria-keyshortcuts={ariaKeyshortcuts} className={`inline-flex min-h-9 items-center gap-1.5 rounded-md px-3 py-1.5 text-sm font-medium transition-colors focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-blue-600 disabled:cursor-not-allowed disabled:opacity-60 ${styles}`}>
      {children}
    </button>
  );
}

export function Status({ kind, children }: { kind: "info" | "error" | "ok"; children: ReactNode }) {
  const styles = {
    info: "border-slate-300 bg-slate-50 text-slate-800 dark:border-slate-600 dark:bg-slate-800 dark:text-slate-200",
    ok: "border-green-300 bg-green-50 text-green-900 dark:border-green-800 dark:bg-green-950 dark:text-green-100",
    error: "border-red-300 bg-red-50 text-red-900 dark:border-red-800 dark:bg-red-950 dark:text-red-100",
  }[kind];
  return (
    <p role={kind === "error" ? "alert" : "status"} className={`rounded-md border px-3 py-2 text-sm ${styles}`}>
      {children}
    </p>
  );
}

export function Fieldset({ legend, children }: { legend: string; children: ReactNode }) {
  return (
    <fieldset className="min-w-0 rounded-md border border-slate-200 p-3 dark:border-slate-700">
      <legend className="px-1 text-sm font-semibold text-slate-800 dark:text-slate-200">{legend}</legend>
      {children}
    </fieldset>
  );
}
