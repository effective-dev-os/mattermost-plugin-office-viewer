// Copyright (c) 2015-present Mattermost, Inc. All Rights Reserved.
// See LICENSE.txt for license information.

import {isOfficeFileInfo, officeFormatOf} from 'office/is_office_file';

import type {FileInfo} from '@mattermost/types/files';

function makeFileInfo(overrides: Partial<FileInfo>): FileInfo {
    return {
        id: 'file1',
        user_id: 'user1',
        channel_id: 'channel1',
        create_at: 0,
        update_at: 0,
        delete_at: 0,
        name: 'document.docx',
        extension: 'docx',
        size: 1024,
        mime_type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        width: 0,
        height: 0,
        has_preview_image: false,
        clientId: 'client1',
        archived: false,
        ...overrides,
    };
}

describe('officeFormatOf / isOfficeFileInfo', () => {
    test.each([
        ['docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'docx'],
        ['DOCX', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'docx'],
        ['pptx', 'application/vnd.openxmlformats-officedocument.presentationml.presentation', 'pptx'],
        ['xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'xlsx'],
        ['jpg', 'image/jpeg', null],
        ['png', 'image/png', null],
    ])('extension=%s mime=%s -> %s', (extension, mimeType, expected) => {
        const fileInfo = makeFileInfo({extension, mime_type: mimeType});
        expect(officeFormatOf(fileInfo)).toBe(expected);
        expect(isOfficeFileInfo(fileInfo)).toBe(expected !== null);
    });

    test('matches by mime type when extension is unrelated', () => {
        const fileInfo = makeFileInfo({
            extension: '',
            mime_type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
        expect(officeFormatOf(fileInfo)).toBe('xlsx');
    });

    test('is case-insensitive for mime type', () => {
        const fileInfo = makeFileInfo({
            extension: '',
            mime_type: 'APPLICATION/VND.OPENXMLFORMATS-OFFICEDOCUMENT.PRESENTATIONML.PRESENTATION',
        });
        expect(officeFormatOf(fileInfo)).toBe('pptx');
    });

    test('returns null when archived, even if extension matches', () => {
        const fileInfo = makeFileInfo({archived: true});
        expect(officeFormatOf(fileInfo)).toBeNull();
        expect(isOfficeFileInfo(fileInfo)).toBe(false);
    });

    test('returns null when the file exceeds the size ceiling', () => {
        const fileInfo = makeFileInfo({size: (25 * 1024 * 1024) + 1});
        expect(officeFormatOf(fileInfo)).toBeNull();
        expect(isOfficeFileInfo(fileInfo)).toBe(false);
    });

    test('allows a file exactly at the size ceiling', () => {
        const fileInfo = makeFileInfo({size: 25 * 1024 * 1024});
        expect(officeFormatOf(fileInfo)).toBe('docx');
    });

    test('returns null for a non-office file with no matching extension or mime', () => {
        const fileInfo = makeFileInfo({extension: 'txt', mime_type: 'text/plain'});
        expect(officeFormatOf(fileInfo)).toBeNull();
        expect(isOfficeFileInfo(fileInfo)).toBe(false);
    });
});
