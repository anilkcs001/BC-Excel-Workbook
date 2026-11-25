// Vendor entry: bundles Univer (Apache-2.0) into a single IIFE exposed as window.AKUniver.
// Rebuild with: npm run build   (output -> ../ControlAddin/scripts/vendor + stylesheets/vendor)
import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets';
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import sheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US';
import { UniverSheetsFilterPreset } from '@univerjs/preset-sheets-filter';
import sheetsFilterEnUS from '@univerjs/preset-sheets-filter/locales/en-US';
import { UniverSheetsSortPreset } from '@univerjs/preset-sheets-sort';
import sheetsSortEnUS from '@univerjs/preset-sheets-sort/locales/en-US';
import { UniverSheetsConditionalFormattingPreset } from '@univerjs/preset-sheets-conditional-formatting';
import sheetsCfEnUS from '@univerjs/preset-sheets-conditional-formatting/locales/en-US';
import { UniverSheetsFindReplacePreset } from '@univerjs/preset-sheets-find-replace';
import sheetsFindEnUS from '@univerjs/preset-sheets-find-replace/locales/en-US';
import { UniverSheetsDataValidationPreset } from '@univerjs/preset-sheets-data-validation';
import sheetsDvEnUS from '@univerjs/preset-sheets-data-validation/locales/en-US';
import { UniverSheetsDrawingPreset } from '@univerjs/preset-sheets-drawing';
import sheetsDrawingEnUS from '@univerjs/preset-sheets-drawing/locales/en-US';
import * as React from 'react';

import '@univerjs/preset-sheets-core/lib/index.css';
import '@univerjs/preset-sheets-filter/lib/index.css';
import '@univerjs/preset-sheets-sort/lib/index.css';
import '@univerjs/preset-sheets-conditional-formatting/lib/index.css';
import '@univerjs/preset-sheets-find-replace/lib/index.css';
import '@univerjs/preset-sheets-data-validation/lib/index.css';
import '@univerjs/preset-sheets-drawing/lib/index.css';

export function create(containerId) {
  return createUniver({
    locale: LocaleType.EN_US,
    locales: {
      [LocaleType.EN_US]: mergeLocales(sheetsCoreEnUS, sheetsFilterEnUS, sheetsSortEnUS, sheetsCfEnUS, sheetsFindEnUS, sheetsDvEnUS, sheetsDrawingEnUS),
    },
    presets: [
      UniverSheetsCorePreset({ container: containerId }),
      UniverSheetsFilterPreset(),
      UniverSheetsSortPreset(),
      UniverSheetsConditionalFormattingPreset(),
      UniverSheetsFindReplacePreset(),
      UniverSheetsDataValidationPreset(),
      // Floating objects over cells (Stage 7 charts and KPI cards): movable, resizable, saved in the snapshot.
      UniverSheetsDrawingPreset(),
    ],
  });
}
// React is exposed so the add-in can register its chart component for the floating objects.
export { LocaleType, React };
