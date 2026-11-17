import { useState, type Ref } from "react";

interface Props {
  accept: string;
  onFile: (file: File) => void;
  inputRef?: Ref<HTMLInputElement>;
  label: string;
  hint: string;
}

/** A file picker that also accepts drag and drop. Files are read in the browser only. */
export function DropZone({ accept, onFile, inputRef, label, hint }: Props) {
  const [over, setOver] = useState(false);
  return (
    <div
      onDragOver={(e) => {
        e.preventDefault();
        setOver(true);
      }}
      onDragLeave={() => setOver(false)}
      onDrop={(e) => {
        e.preventDefault();
        setOver(false);
        const f = e.dataTransfer.files[0];
        if (f) onFile(f);
      }}
      className={`flex flex-col items-center justify-center gap-2 rounded-lg border-2 border-dashed px-4 py-6 text-center transition-colors ${over ? "border-blue-600 bg-blue-50 dark:bg-blue-950" : "border-slate-300 dark:border-slate-600"}`}
    >
      <label className="cursor-pointer rounded-md bg-blue-700 px-4 py-2 text-sm font-medium text-white hover:bg-blue-800 focus-within:outline-2 focus-within:outline-offset-2 focus-within:outline-blue-600">
        {label}
        <input
          ref={inputRef}
          type="file"
          accept={accept}
          className="sr-only"
          onChange={(e) => {
            const f = e.target.files?.[0];
            if (f) onFile(f);
            e.target.value = "";
          }}
        />
      </label>
      <p className="text-sm text-slate-700 dark:text-slate-300">{hint}</p>
    </div>
  );
}
