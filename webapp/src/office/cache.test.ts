// Copyright (c) 2015-present Mattermost, Inc. All Rights Reserved.
// See LICENSE.txt for license information.

import type * as CacheModule from 'office/cache';

let cache: typeof CacheModule;

const MB = 1024 * 1024;

beforeEach(() => {
    jest.resetModules();

    // The module-scoped cache is a singleton keyed by import, so each test
    // needs a fresh module instance to avoid state leaking across tests.
    // eslint-disable-next-line global-require
    cache = require('office/cache');
});

describe('office file buffer cache', () => {
    test('returns undefined for an uncached file id', () => {
        expect(cache.getCachedFileBuffer('missing')).toBeUndefined();
    });

    test('returns a previously cached buffer', () => {
        const buffer = new ArrayBuffer(10);
        cache.setCachedFileBuffer('file1', buffer);
        expect(cache.getCachedFileBuffer('file1')).toBe(buffer);
    });

    test('overwriting the same file id replaces the cached buffer', () => {
        const first = new ArrayBuffer(10);
        const second = new ArrayBuffer(20);
        cache.setCachedFileBuffer('file1', first);
        cache.setCachedFileBuffer('file1', second);

        expect(cache.getCachedFileBuffer('file1')).toBe(second);
    });

    test('evicts the least-recently-used entry once the byte budget is exceeded', () => {
        // Budget is 64MB. Three 25MB entries (75MB total) exceed it, so the
        // oldest (file0) must be evicted to bring the total back at or under
        // budget.
        cache.setCachedFileBuffer('file0', new ArrayBuffer(25 * MB));
        cache.setCachedFileBuffer('file1', new ArrayBuffer(25 * MB));
        cache.setCachedFileBuffer('file2', new ArrayBuffer(25 * MB));

        expect(cache.getCachedFileBuffer('file0')).toBeUndefined();
        expect(cache.getCachedFileBuffer('file1')).toBeDefined();
        expect(cache.getCachedFileBuffer('file2')).toBeDefined();
    });

    test('accessing an entry marks it as most-recently-used, protecting it from eviction', () => {
        cache.setCachedFileBuffer('file0', new ArrayBuffer(25 * MB));
        cache.setCachedFileBuffer('file1', new ArrayBuffer(25 * MB));

        // Touch file0 so it becomes most-recently-used.
        cache.getCachedFileBuffer('file0');

        // Adding one more 25MB entry pushes the total to 75MB, over the
        // 64MB budget, so the new LRU (file1) should be evicted instead of
        // file0.
        cache.setCachedFileBuffer('file2', new ArrayBuffer(25 * MB));

        expect(cache.getCachedFileBuffer('file0')).toBeDefined();
        expect(cache.getCachedFileBuffer('file1')).toBeUndefined();
        expect(cache.getCachedFileBuffer('file2')).toBeDefined();
    });
});
