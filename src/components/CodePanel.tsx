// The code that reproduces the current settings, with install instructions.
import { useState } from "react";
import { installCommand, snippetText, type PackageManager, type Snippet } from "../lib/codegen";
import { Button, Card } from "./ui";

const PMS: PackageManager[] = ["npm", "pnpm", "yarn", "bun"];

function CopyButton({ text, label }: { text: string; label: string }) {
  const [done, setDone] = useState(false);
  return (
    <Button
      variant="ghost"
      onClick={() => {
        navigator.clipboard.writeText(text).then(
          () => {
            setDone(true);
            setTimeout(() => setDone(false), 1500);
          },
          () => setDone(false),
        );
      }}
    >
      <span aria-live="polite">{done ? "Copied" : label}</span>
    </Button>
  );
}

export function CodePanel({ snippet, version }: { snippet: Snippet; version: string }) {
  const [pm, setPm] = useState<PackageManager>("npm");
  const install = installCommand(pm, version);
  return (
    <Card title="Code" actions={<CopyButton text={snippetText(snippet)} label="Copy all" />}>
      <div className="space-y-3">
        <div>
          <div role="radiogroup" aria-label="Package manager" className="mb-2 flex flex-wrap gap-1">
            {PMS.map((p) => (
              <button
                key={p}
                type="button"
                role="radio"
                aria-checked={p === pm}
                onClick={() => setPm(p)}
                className={`rounded px-2 py-0.5 text-xs font-medium ${p === pm ? "bg-slate-800 text-white dark:bg-slate-200 dark:text-slate-900" : "text-slate-700 hover:bg-slate-100 dark:text-slate-300 dark:hover:bg-slate-800"}`}
              >
                {p}
              </button>
            ))}
          </div>
          <div className="flex items-center gap-2">
            <pre className="code flex-1 overflow-x-auto">{install}</pre>
            <CopyButton text={install} label="Copy" />
          </div>
          <p className="mt-1 text-xs text-slate-600 dark:text-slate-400">
            {version} is a prerelease on the npm <code>next</code> tag; pin the exact version.
          </p>
        </div>
        {snippet.files.map((f) => (
          <div key={f.name}>
            <div className="mb-1 flex items-center justify-between">
              <span className="font-mono text-xs text-slate-700 dark:text-slate-300">{f.name}</span>
              <CopyButton text={f.code} label="Copy" />
            </div>
            <pre className="code max-h-96 overflow-auto">{f.code}</pre>
          </div>
        ))}
      </div>
    </Card>
  );
}
