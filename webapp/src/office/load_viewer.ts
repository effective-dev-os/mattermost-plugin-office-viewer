// Copyright (c) 2015-present Mattermost, Inc. All Rights Reserved.
// See LICENSE.txt for license information.

import type {OfficeFormat} from 'office/is_office_file';

// This is the ONLY module allowed to import from '@silurus/ooxml'. Every
// entry point goes behind await import('@silurus/ooxml/<format>') so the
// library (WASM parser + render-worker chunk + renderer code) is fetched
// only when a matching file is actually opened, never eagerly bundled into
// main.js. A static top-level import anywhere reachable from index.tsx
// would balloon main.js from ~3KB to 1.4MB+.
//
// Every viewer below is constructed with mode: 'worker'. @silurus/ooxml's
// default mode ('main') builds its parser Worker via a Blob +
// URL.createObjectURL() (with a data: URL fallback) — both are blocked by
// Mattermost's CSP (script-src 'self', no worker-src/blob:/data:).
// mode: 'worker' instead builds the Worker from a real emitted same-origin
// chunk file via new URL(asset, import.meta.url), which CSP allows. This is
// not a style preference — mode: 'main' will not work in Mattermost.

// Untrusted, client-side-parsed input (an uploaded .docx/.pptx/.xlsx):
// bound resource use to guard against a resource-exhaustion DoS in the
// viewer's tab. WASM sandboxing already contains memory-safety; this is
// about bounding CPU/memory consumption, not RCE.
const RESOURCE_LIMITS = {
    maxArchiveEntryBytes: 64 * 1024 * 1024,
    maxTotalInflatedBytes: 128 * 1024 * 1024,
    maxArchiveEntries: 2048,
};
const MAX_ZIP_ENTRY_BYTES = 64 * 1024 * 1024;
const WORKER_TIMEOUT_MS = 30000;

export type OfficeViewerHandle = {
    destroy: () => void;
};

export async function loadOfficeViewer(
    format: OfficeFormat,
    container: HTMLElement,
    buffer: ArrayBuffer,
    onError: (err: Error) => void,
): Promise<OfficeViewerHandle> {
    switch (format) {
    case 'docx': {
        const {DocxScrollViewer} = await import('@silurus/ooxml/docx');
        const viewer = new DocxScrollViewer(container, {
            mode: 'worker',
            maxZipEntryBytes: MAX_ZIP_ENTRY_BYTES,
            resourceLimits: RESOURCE_LIMITS,
            workerTimeoutMs: WORKER_TIMEOUT_MS,
            onError,
        });
        await viewer.load(buffer);
        return {destroy: () => viewer.destroy()};
    }
    case 'pptx': {
        const {PptxScrollViewer} = await import('@silurus/ooxml/pptx');
        const viewer = new PptxScrollViewer(container, {
            mode: 'worker',
            maxZipEntryBytes: MAX_ZIP_ENTRY_BYTES,
            resourceLimits: RESOURCE_LIMITS,
            workerTimeoutMs: WORKER_TIMEOUT_MS,
            onError,
        });
        await viewer.load(buffer);
        return {destroy: () => viewer.destroy()};
    }
    case 'xlsx': {
        const {XlsxViewer} = await import('@silurus/ooxml/xlsx');
        const viewer = new XlsxViewer(container, {
            mode: 'worker',
            maxZipEntryBytes: MAX_ZIP_ENTRY_BYTES,
            resourceLimits: RESOURCE_LIMITS,
            workerTimeoutMs: WORKER_TIMEOUT_MS,
            onError,
        });
        await viewer.load(buffer);
        return {destroy: () => viewer.destroy()};
    }
    default: {
        const exhaustive: never = format;
        throw new Error(`Unsupported office format: ${String(exhaustive)}`);
    }
    }
}
