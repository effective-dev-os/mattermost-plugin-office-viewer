// Copyright (c) 2015-present Mattermost, Inc. All Rights Reserved.
// See LICENSE.txt for license information.

// @silurus/ooxml only exposes its docx/pptx/xlsx entry points through the
// package.json "exports" map — there is no physical docx.d.ts/pptx.d.ts/
// xlsx.d.ts at the package root, and no "typesVersions" fallback. This
// repo's tsconfig.json uses classic Node module resolution
// ("moduleResolution": "node", TypeScript 4.9.5), which does not honor the
// "exports" map for type resolution and so cannot statically resolve
// '@silurus/ooxml/docx' et al. — even though webpack 5 (which does honor
// "exports") resolves the same specifiers fine at build/runtime.
//
// These ambient declarations describe only the surface load_viewer.ts
// actually uses. Verified by hand against the installed package's own
// node_modules/@silurus/ooxml/dist/types/{docx,pptx,xlsx}.d.ts.

type SilurusOoxmlResourceLimits = {
    maxArchiveEntryBytes?: number | null;
    maxTotalInflatedBytes?: number | null;
    maxArchiveEntries?: number | null;
};

type SilurusOoxmlLoadOptions = {
    mode?: 'main' | 'worker';
    maxZipEntryBytes?: number;
    resourceLimits?: SilurusOoxmlResourceLimits;
    workerTimeoutMs?: number;
    onError?: (err: Error) => void;
};

declare module '@silurus/ooxml/docx' {
    export type DocxScrollViewerOptions = SilurusOoxmlLoadOptions;
    export class DocxScrollViewer {
        constructor(container: HTMLElement, opts?: DocxScrollViewerOptions);
        load(source: string | ArrayBuffer): Promise<void>;
        destroy(): void;
    }
}

declare module '@silurus/ooxml/pptx' {
    export type PptxScrollViewerOptions = SilurusOoxmlLoadOptions;
    export class PptxScrollViewer {
        constructor(container: HTMLElement, opts?: PptxScrollViewerOptions);
        load(source: string | ArrayBuffer): Promise<void>;
        destroy(): void;
    }
}

declare module '@silurus/ooxml/xlsx' {
    export type XlsxViewerOptions = SilurusOoxmlLoadOptions;
    export class XlsxViewer {
        constructor(container: HTMLElement, opts?: XlsxViewerOptions);
        load(source: string | ArrayBuffer): Promise<void>;
        destroy(): void;
    }
}
