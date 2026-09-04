// Copyright (c) 2015-present Mattermost, Inc. All Rights Reserved.
// See LICENSE.txt for license information.

// Cache of the raw fetched ArrayBuffer, keyed by fileInfo.id — NOT a parsed
// engine/viewer instance. Those own a live Worker + WASM heap released only
// by their own destroy(); caching them would pin memory and leak workers
// across scroll/reopen. Bytes are cheap to hold and cheap to re-parse from.
const MAX_CACHED_BYTES = 64 * 1024 * 1024;

// Map preserves insertion order; re-inserting a key on access moves it to
// the end, giving us LRU ordering for free without a separate structure.
const buffersByFileId = new Map<string, ArrayBuffer>();
let totalCachedBytes = 0;

export function getCachedFileBuffer(fileId: string): ArrayBuffer | undefined {
    const buffer = buffersByFileId.get(fileId);
    if (buffer === undefined) {
        return undefined;
    }

    buffersByFileId.delete(fileId);
    buffersByFileId.set(fileId, buffer);
    return buffer;
}

export function setCachedFileBuffer(fileId: string, buffer: ArrayBuffer): void {
    const existing = buffersByFileId.get(fileId);
    if (existing !== undefined) {
        totalCachedBytes -= existing.byteLength;
        buffersByFileId.delete(fileId);
    }

    buffersByFileId.set(fileId, buffer);
    totalCachedBytes += buffer.byteLength;

    while (totalCachedBytes > MAX_CACHED_BYTES) {
        const oldestFileId = buffersByFileId.keys().next().value;
        if (oldestFileId === undefined) {
            break;
        }
        const oldestBuffer = buffersByFileId.get(oldestFileId);
        buffersByFileId.delete(oldestFileId);
        if (oldestBuffer !== undefined) {
            totalCachedBytes -= oldestBuffer.byteLength;
        }
    }
}
