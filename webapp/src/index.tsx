// Copyright (c) 2015-present Mattermost, Inc. All Rights Reserved.
// See LICENSE.txt for license information.

import manifest from 'manifest';
import {isOfficeFileInfo} from 'office/is_office_file';
import type {Store} from 'redux';

import type {GlobalState} from '@mattermost/types/store';

import {OfficeFilePreview} from 'components/office_file_preview';

import type {PluginRegistry} from 'types/mattermost-webapp';

export default class Plugin {
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    public async initialize(registry: PluginRegistry, store: Store<GlobalState>) {
        registry.registerFilePreviewComponent(
            (fileInfo) => isOfficeFileInfo(fileInfo),
            OfficeFilePreview,
        );
    }
}

declare global {
    interface Window {
        registerPlugin(pluginId: string, plugin: Plugin): void;
    }
}

window.registerPlugin(manifest.id, new Plugin());
