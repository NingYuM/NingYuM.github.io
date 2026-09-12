// Web Worker for WASM PPTX parsing — runs decompress/parse/serialize off the main thread.

import { arrayBufferToLatin1, uint8ArrayToLatin1 } from './latin1'
import { createWasmStringConstants, fetchWasmBytes } from './wasm-loader'
import {
  checkBrowserElapsed,
  checkBrowserInput,
  resourcePolicyErrorFromDetails,
  unzipBrowserPptx,
} from './resource-policy'
import { serializeWorkerError } from './worker-error'

interface WasmExports {
  parse_pptx_from_string: (data: string) => string
  decompress: (data: string) => number
  parse_only: () => number
  serialize_result: () => string
  parse_decompressed: () => string
  add_predecompressed_file: (filename: string, content_latin1: string) => number
  parse_from_predecompressed: () => string
  _start: () => void
}

let instance: WasmExports | null = null
let instanceSource: string | null = null

interface WorkerRequest {
  id?: number
  buffer: ArrayBuffer
  unzipMode: 'js' | 'wasm'
  wasmUrl?: string
}

async function ensureWasm(wasmUrl?: string): Promise<WasmExports> {
  const sourceUrl = wasmUrl ?? new URL('./mbt/main.wasm', import.meta.url)
  const source = String(sourceUrl)
  if (instance && instanceSource === source) return instance
  const wasmBytes = await fetchWasmBytes(sourceUrl, 'wasm-worker')
  const { instance: inst } = (await WebAssembly.instantiate(
    new Uint8Array(wasmBytes),
    { _: createWasmStringConstants() },
    // @ts-expect-error -- wasm-gc builtins API not yet in TS lib types
    { builtins: ['js-string'], importedStringConstants: '_' },
  )) as WebAssembly.WebAssemblyInstantiatedSource
  instance = inst.exports as unknown as WasmExports
  instanceSource = source
  instance._start()
  return instance
}

async function parseWithWasmUnzip(mod: WasmExports, buffer: ArrayBuffer) {
  const t0 = performance.now()
  checkBrowserInput(buffer.byteLength)
  const latin1 = arrayBufferToLatin1(buffer)
  const t1 = performance.now()

  const fileCount = mod.decompress(latin1)
  const t2 = performance.now()

  mod.parse_only()
  const t3 = performance.now()

  const jsonStr = mod.serialize_result()
  const t4 = performance.now()

  const result = JSON.parse(jsonStr)
  const t5 = performance.now()

  const violation = resourcePolicyErrorFromDetails(result.resourceViolation)
  if (violation) throw violation
  checkBrowserElapsed(t0, '[wasm-parser] Worker parse')
  return {
    result,
    timing: {
      latin1: t1 - t0,
      decompress: t2 - t1,
      fileCount,
      parse: t3 - t2,
      serialize: t4 - t3,
      jsonParse: t5 - t4,
      total: t5 - t0,
    },
  }
}

async function parseWithJsUnzip(mod: WasmExports, buffer: ArrayBuffer) {
  const t0 = performance.now()
  checkBrowserInput(buffer.byteLength)
  const files = unzipBrowserPptx(buffer)
  const t1 = performance.now()

  let fileCount = 0
  for (const [filename, data] of Object.entries(files)) {
    mod.add_predecompressed_file(filename, uint8ArrayToLatin1(data))
    fileCount++
  }
  const t2 = performance.now()

  const jsonStr = mod.parse_from_predecompressed()
  const t3 = performance.now()

  const result = JSON.parse(jsonStr)
  const t4 = performance.now()

  const violation = resourcePolicyErrorFromDetails(result.resourceViolation)
  if (violation) throw violation
  checkBrowserElapsed(t0, '[wasm-parser] Worker parse')
  return {
    result,
    timing: {
      latin1: 0,
      decompress: t1 - t0,
      fileCount,
      parse: t3 - t2,
      serialize: t2 - t1,
      jsonParse: t4 - t3,
      total: t4 - t0,
    },
  }
}

const active = new Set<number>()

self.onmessage = async (
  e: MessageEvent<WorkerRequest | ArrayBuffer | { kind: 'cancel'; id: number }>,
) => {
  if ('kind' in e.data && e.data.kind === 'cancel') {
    active.delete(e.data.id)
    return
  }
  const id = e.data instanceof ArrayBuffer ? 0 : (e.data.id ?? 0)
  active.add(id)
  try {
    const request =
      e.data instanceof ArrayBuffer
        ? { buffer: e.data, unzipMode: 'wasm' as const }
        : (e.data as WorkerRequest)
    const { buffer, unzipMode, wasmUrl } = request
    const mod = await ensureWasm(wasmUrl)
    if (!active.has(id)) return

    const parsed =
      unzipMode === 'js'
        ? await parseWithJsUnzip(mod, buffer)
        : await parseWithWasmUnzip(mod, buffer)
    const { result, timing } = parsed
    if (!active.has(id)) return

    if (result.error) {
      self.postMessage({
        id,
        status: 'error',
        error: serializeWorkerError(new Error(`WASM parser error: ${result.error}`)),
      })
      return
    }

    self.postMessage({
      id,
      status: 'ok',
      data: result,
      timing,
    })
  } catch (err) {
    if (!active.has(id)) return
    self.postMessage({
      id,
      status: 'error',
      error: serializeWorkerError(err),
    })
  } finally {
    // A cancel message removes the active id; acknowledge only after pending
    // initialization or parsing exits so the host can safely retire its deadline.
    if (!active.delete(id)) {
      self.postMessage({
        id,
        status: 'error',
        error: serializeWorkerError(new DOMException('The operation was aborted.', 'AbortError')),
      })
    }
  }
}
