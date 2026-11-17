// The playground's Web Worker: runs the reader, writer and CSV code off the main thread.
import { exposeWorker } from "javascript-excel/worker";

exposeWorker();
