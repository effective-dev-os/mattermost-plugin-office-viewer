// Copyright (c) 2015-present Mattermost, Inc. All Rights Reserved.
// See LICENSE.txt for license information.

import {useOfficePreview} from 'office/use_office_preview';
import React, {useRef} from 'react';

import type {FileInfo} from '@mattermost/types/files';
import type {Post} from '@mattermost/types/posts';

const TEXT_COLOR = 'rgba(255, 255, 255, 0.88)';
const PREVIEW_HEIGHT = 'calc(100vh - 168px)';

type OfficeFilePreviewProps = {
    fileInfo: FileInfo;
    post?: Post;
    onModalDismissed: () => void;
};

function OfficeFallback({fileInfo, reason}: {fileInfo: FileInfo; reason?: string}) {
    return (
        <div style={{padding: '20px', textAlign: 'center', color: TEXT_COLOR}}>
            {reason && <p>{reason}</p>}
            <p>{fileInfo.name}</p>
            <a
                href={`/api/v4/files/${fileInfo.id}?download=1`}
                download={fileInfo.name}
            >
                {'Download'}
            </a>
        </div>
    );
}

function OfficePreviewBody({fileInfo}: {fileInfo: FileInfo}) {
    const containerRef = useRef<HTMLDivElement>(null);
    const officePreview = useOfficePreview(fileInfo, containerRef);

    if (officePreview.status === 'failed') {
        // eslint-disable-next-line no-console
        console.error('Office preview failed to render:', officePreview.error);
    }

    return (
        <div style={{position: 'relative', width: '100%', height: PREVIEW_HEIGHT}}>
            <div
                ref={containerRef}
                style={{
                    width: '100%',
                    height: PREVIEW_HEIGHT,
                    visibility: officePreview.status === 'ready' ? 'visible' : 'hidden',
                }}
            />
            {officePreview.status === 'loading' && (
                <div
                    role='status'
                    aria-live='polite'
                    style={{
                        position: 'absolute',
                        top: 0,
                        left: 0,
                        width: '100%',
                        height: '100%',
                        padding: '20px',
                        textAlign: 'center',
                        color: TEXT_COLOR,
                    }}
                >
                    {'Loading document preview…'}
                </div>
            )}
            {officePreview.status === 'failed' && (
                <div
                    style={{
                        position: 'absolute',
                        top: 0,
                        left: 0,
                        width: '100%',
                        height: '100%',
                    }}
                >
                    <OfficeFallback
                        fileInfo={fileInfo}
                        reason={'This document could not be displayed.'}
                    />
                </div>
            )}
        </div>
    );
}

type OfficeFilePreviewState = {
    hasError: boolean;
};

// React error boundaries only work as class components (no hook
// equivalent) — a render exception here must never crash Mattermost's
// post-render tree.
export class OfficeFilePreview extends React.Component<OfficeFilePreviewProps, OfficeFilePreviewState> {
    constructor(props: OfficeFilePreviewProps) {
        super(props);
        this.state = {hasError: false};
    }

    static getDerivedStateFromError(): OfficeFilePreviewState {
        return {hasError: true};
    }

    componentDidUpdate(prevProps: OfficeFilePreviewProps) {
        if (prevProps.fileInfo.id !== this.props.fileInfo.id && this.state.hasError) {
            this.setState({hasError: false});
        }
    }

    render() {
        if (this.state.hasError) {
            return <OfficeFallback fileInfo={this.props.fileInfo}/>;
        }

        return (
            <OfficePreviewBody
                key={this.props.fileInfo.id}
                fileInfo={this.props.fileInfo}
            />
        );
    }
}
