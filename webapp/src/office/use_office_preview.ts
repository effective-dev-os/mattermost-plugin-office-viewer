// Copyright (c) 2015-present Mattermost, Inc. All Rights Reserved.
// See LICENSE.txt for license information.

import {getCachedFileBuffer, setCachedFileBuffer} from 'office/cache';
import {officeFormatOf} from 'office/is_office_file';
import type {OfficeViewerHandle} from 'office/load_viewer';
import {loadOfficeViewer} from 'office/load_viewer';
import {useEffect, useState} from 'react';

import type {FileInfo} from '@mattermost/types/files';

export type OfficePreviewState =
    | {status: 'loading'}
    | {status: 'ready'}
    | {status: 'failed'; error: string};

export function useOfficePreview(fileInfo: FileInfo, containerRef: React.RefObject<HTMLDivElement>): OfficePreviewState {
    const [state, setState] = useState<OfficePreviewState>({status: 'loading'});

    useEffect(() => {
        let cancelled = false;
        let viewerHandle: OfficeViewerHandle | undefined;

        const format = officeFormatOf(fileInfo);
        if (format === null) {
            setState({status: 'failed', error: 'Unsupported office document format'});
            return () => {
                cancelled = true;
            };
        }

        setState({status: 'loading'});

        (async () => {
            try {
                let buffer = getCachedFileBuffer(fileInfo.id);
                if (buffer === undefined) {
                    const response = await fetch(`/api/v4/files/${fileInfo.id}`, {credentials: 'same-origin'});
                    if (!response.ok) {
                        throw new Error(`Failed to fetch file (status ${response.status})`);
                    }
                    const fetched = await response.arrayBuffer();
                    if (cancelled) {
                        return;
                    }

                    // Cache the pristine, unsliced buffer — only copies of
                    // it are ever handed to the viewer.
                    setCachedFileBuffer(fileInfo.id, fetched);
                    buffer = fetched;
                }

                if (cancelled) {
                    return;
                }

                const container = containerRef.current;
                if (container === null) {
                    throw new Error('Office preview container is not mounted');
                }

                // mode: 'worker' most likely postMessages the bytes to its
                // worker with a transfer list (zero-copy), which detaches
                // the ArrayBuffer object passed in. Pass a copy every time
                // so the cached original stays valid for the next open.
                const handle = await loadOfficeViewer(format, container, buffer.slice(0), (err) => {
                    if (!cancelled) {
                        setState({status: 'failed', error: err.message});
                    }
                });

                if (cancelled) {
                    handle.destroy();
                    return;
                }

                viewerHandle = handle;
                setState({status: 'ready'});
            } catch (err) {
                if (!cancelled) {
                    setState({status: 'failed', error: err instanceof Error ? err.message : 'Unknown error'});
                }
            }
        })();

        return () => {
            cancelled = true;
            viewerHandle?.destroy();
        };
    }, [fileInfo.id]);

    return state;
}
