// Copyright (c) 2015-present Mattermost, Inc. All Rights Reserved.
// See LICENSE.txt for license information.

import type {FileInfo} from '@mattermost/types/files';

export type OfficeFormat = 'docx' | 'pptx' | 'xlsx';

// @silurus/ooxml parses the whole archive client-side in the tab; past
// this size, let Mattermost's default download UI handle it instead of
// tying up the WASM heap for an unrequested large-file preview.
const MAX_OFFICE_FILE_BYTES = 25 * 1024 * 1024;

const OFFICE_EXTENSIONS: Record<string, OfficeFormat> = {
    docx: 'docx',
    pptx: 'pptx',
    xlsx: 'xlsx',
};

// OOXML MIME types ARE in Go's builtin mime.TypeByExtension table (unlike
// HEIC), but server MIME still depends on the Go version / /etc/mime.types
// override, so extension stays the primary signal and mime is a fallback.
const OFFICE_MIME_TYPES: Record<string, OfficeFormat> = {
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'docx',
    'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'pptx',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'xlsx',
};

export function officeFormatOf(fileInfo: FileInfo): OfficeFormat | null {
    if (fileInfo.archived) {
        return null;
    }

    if (fileInfo.size > MAX_OFFICE_FILE_BYTES) {
        return null;
    }

    const byExtension = OFFICE_EXTENSIONS[fileInfo.extension.toLowerCase()];
    if (byExtension) {
        return byExtension;
    }

    return OFFICE_MIME_TYPES[fileInfo.mime_type.toLowerCase()] ?? null;
}

export function isOfficeFileInfo(fileInfo: FileInfo): boolean {
    return officeFormatOf(fileInfo) !== null;
}
