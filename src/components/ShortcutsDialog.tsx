import { useEffect, useRef } from "react";

export interface Shortcut {
  keys: string;
  description: string;
}

export function ShortcutsDialog({ open, onClose, shortcuts }: { open: boolean; onClose: () => void; shortcuts: Shortcut[] }) {
  const ref = useRef<HTMLDialogElement>(null);
  useEffect(() => {
    const d = ref.current;
    if (!d) return;
    if (open && !d.open) d.showModal();
    if (!open && d.open) d.close();
  }, [open]);
  return (
    <dialog
      ref={ref}
      onClose={onClose}
      onClick={(e) => {
        if (e.target === ref.current) onClose();
      }}
      aria-labelledby="shortcuts-title"
      className="m-auto w-[min(32rem,calc(100vw-2rem))] rounded-lg bg-white p-0 text-slate-900 shadow-xl backdrop:bg-black/50 dark:bg-slate-900 dark:text-slate-100"
    >
      <div className="p-5">
        <h2 id="shortcuts-title" className="mb-3 text-lg font-semibold">
          Keyboard shortcuts
        </h2>
        <table className="w-full text-sm">
          <tbody>
            {shortcuts.map((s) => (
              <tr key={s.keys} className="border-t border-slate-200 dark:border-slate-700">
                <td className="py-1.5 pr-3">
                  <kbd className="rounded border border-slate-300 bg-slate-100 px-1.5 py-0.5 font-mono text-xs dark:border-slate-600 dark:bg-slate-800">{s.keys}</kbd>
                </td>
                <td className="py-1.5">{s.description}</td>
              </tr>
            ))}
          </tbody>
        </table>
        <div className="mt-4 text-right">
          <button type="button" onClick={onClose} className="rounded-md bg-blue-700 px-3 py-1.5 text-sm font-medium text-white hover:bg-blue-800">
            Close
          </button>
        </div>
      </div>
    </dialog>
  );
}
