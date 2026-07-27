# Working With HTML Examples

This folder contains **Aspose.Cells for .NET** code examples related to:

Working With HTML


## Purpose

These examples demonstrate common **Aspose.Cells APIs** used when working with:

- Workbooks
- Worksheets
- Cells
- Formulas
- Charts
- Data operations


## Example Files

Each `.cs` file demonstrates a specific task related to **Working With HTML**.

Example:

create-a-workbook.cs


## Required Namespaces

Most examples will require:

using Aspose.Cells;


## Common Pattern

Typical Aspose.Cells workflow:

Workbook workbook = new Workbook();

Worksheet sheet = workbook.Worksheets[0];

Cells cells = sheet.Cells;


## Output

Examples may generate:

- XLSX files
- PDF files
- CSV files
- Images

Output files are written to the working directory.
- verify-that-html-output-contains-only-standard-comment-tags-without-downlevel-revealed-format.cs
- test-html-files-in-legacy-browsers-to-confirm-comments-are-hidden-appropriately.cs
- document-the-configuration-required-to-turn-off-downlevel-revealed-comments.cs
- compare-page-source-before-and-after-disabling-downlevel-revealed-comments.cs
- apply-this-setting-when-exporting-workbooks-containing-extensive-cell-comments.cs
- ensure-that-comment-visibility-settings-respect-user-privacy-requirements-strictly.cs
- validate-that-disabling-downlevel-revealed-comments-does-not-affect-other-html-features.cs
- load-an-xlsx-workbook-enable-excludeunusedstyles-and-save-as-html-to-reduce-file-size.cs
- load-an-xls-workbook-set-isexportcomments-to-true-and-generate-html-that-includes-all-cell-comments.cs
- load-a-workbook-enable-exportgridlines-and-produce-html-that-displays-worksheet-gridlines.cs
- load-a-workbook-with-conditional-formatting-activate-databar-export-and-save-html-preserving-databar-visuals.cs
- combine-excludeunusedstyles-and-isexportcomments-options-then-save-workbook-as-html-for-compact-output-with-comments.cs
- process-multiple-xlsx-files-in-a-directory-applying-excludeunusedstyles-to-each-and-batch-save-reduced-size-html-files.cs
- iterate-through-all-worksheets-in-a-workbook-saving-each-as-separate-html-files-while-preserving-gridlines.cs
- validate-that-generated-html-contains-comment-tags-by-searching-for-after-enabling-isexportcomments.cs
- compare-html-file-sizes-with-and-without-excludeunusedstyles-to-quantify-reduction-achieved-by-style-omission.cs
- verify-that-unused-named-styles-are-omitted-from-the-html-output-when-excludeunusedstyles-is-enabled.cs
- load-a-workbook-with-hidden-rows-enable-exportgridlines-and-ensure-hidden-rows-do-not-affect-gridline-rendering.cs
- create-a-workbook-programmatically-add-a-comment-to-a-cell-enable-isexportcomments-and-verify-comment-appears-in-html.cs
- export-a-workbook-with-conditional-formatting-that-includes-overlapping-rules-and-ensure-both-rules-are-represented-in-html.cs
- save-a-workbook-to-html-with-exportgridlines-disabled-then-confirm-that-gridlines-are-absent-in-the-output.cs
- apply-a-custom-style-to-a-range-enable-excludeunusedstyles-and-verify-only-used-styles-appear-in-html.cs
- create-a-workbook-with-conditional-formatting-using-iconset-disable-iconset-export-and-confirm-icons-are-omitted-in-html.cs
- batch-convert-a-set-of-xls-files-to-html-applying-exportgridlines-only-to-worksheets-that-contain-tables.cs
- generate-html-from-a-workbook-while-disabling-all-conditional-formatting-export-to-produce-a-plain-visual-representation.cs
- export-a-workbook-with-conditional-formatting-that-includes-databar-and-extract-the-generated-css-classes-for-analysis.cs
- export-a-workbook-with-conditional-formatting-that-uses-threecolor-scales-and-verify-gradient-colors-appear-in-html.cs
- batch-process-workbooks-applying-both-excludeunusedstyles-and-exportgridlines-and-compare-processing-time-against-default-settings.cs
- batch-process-a-folder-of-xlsx-files-applying-excludeunusedstyles-only-to-files-larger-than-5-mb.cs
- export-a-workbook-with-conditional-formatting-that-includes-databar-and-verify-bar-lengths-match-cell-values-in-html.cs
- generate-html-from-a-workbook-while-disabling-excludeunusedstyles-to-compare-full-style-set-against-reduced-version.cs
- batch-process-a-collection-of-workbooks-applying-exportgridlines-and-isexportcomments-simultaneously-to-each-html-output.cs
- export-a-workbook-with-conditional-formatting-that-includes-colorscale-and-extract-the-css-gradient-definitions-for-reuse.cs
- export-a-workbook-with-conditional-formatting-that-includes-iconset-and-verify-icons-are-rendered-using-img-tags.cs
- batch-convert-workbooks-toggling-exportgridlines-based-on-whether-the-source-worksheet-contains-borders.cs
- export-a-workbook-with-conditional-formatting-that-includes-databar-and-generate-a-separate-css-file-containing-bar-styles.cs
- export-a-workbook-with-conditional-formatting-that-includes-colorscale-and-embed-the-scale-definitions-within-a-style-block.cs
- load-an-xlsx-workbook-and-export-to-html-using-default-settings-preserving-all-content.cs
- export-a-workbook-to-html-while-omitting-document-properties-by-setting-exportdocumentproperties-to-false.cs
- export-a-workbook-to-html-while-omitting-workbook-properties-by-disabling-exportworkbookproperties-flag.cs
- export-a-workbook-to-html-while-omitting-worksheet-properties-by-setting-exportworksheetproperties-to-false.cs
- export-a-workbook-to-html-while-omitting-document-workbook-and-worksheet-properties-simultaneously.cs
- export-a-workbook-to-html-while-omitting-workbook-and-worksheet-properties-together.cs
- export-a-workbook-to-html-while-omitting-document-and-worksheet-properties-together.cs
- export-right-to-left-aligned-text-correctly-using-default-htmlsaveoptions-during-conversion.cs
- explicitly-enable-right-to-left-text-support-by-setting-rtltextsupport-to-true-before-saving.cs
- export-only-the-defined-print-area-to-html-by-setting-exportprintareaonly-to-true.cs
- export-defined-print-area-while-omitting-document-properties-by-combining-exportprintareaonly-and-exportdocumentproperties-false.cs
- export-defined-print-area-while-omitting-worksheet-properties-by-combining-exportprintareaonly-and-exportworksheetproperties-false.cs
- export-defined-print-area-while-disabling-css-generation-by-setting-exportprintareaonly-true-and-disablecss-true.cs
- export-a-workbook-to-html-with-both-css-disabled-and-custom-properties-enabled.cs
- export-a-workbook-to-html-with-custom-properties-enabled-while-preserving-default-css-generation.cs
- disable-css-generation-during-html-export-by-setting-disablecss-to-true.cs
- enable-custom-css-properties-during-html-export-by-setting-enablecustomproperties-to-true.cs
- load-an-html-file-into-a-workbook-while-preserving-div-tag-layout-using-enabledivtaglayout.cs
- load-an-html-file-into-a-workbook-without-preserving-div-layout-by-leaving-enabledivtaglayout-false.cs
- load-html-preserving-div-layout-then-re-export-to-html-and-verify-div-structure-remains-unchanged.cs
- load-html-modify-a-cell-value-programmatically-and-export-to-html-with-default-options.cs
- load-html-modify-a-cell-and-export-to-html-while-disabling-css-generation.cs
- load-html-preserving-div-layout-modify-a-cell-and-export-to-html-while-disabling-css-generation.cs
- batch-convert-all-xlsx-files-in-a-directory-to-html-using-default-conversion-settings.cs
- batch-convert-xlsx-files-to-html-while-omitting-document-properties-for-privacy-compliance.cs
- batch-convert-xlsx-files-to-html-while-disabling-css-generation-to-reduce-file-size.cs
- export-workbook-to-html-after-defining-a-print-area-ensuring-only-that-area-appears.cs
- export-workbook-to-html-after-defining-a-print-area-and-omitting-worksheet-properties.cs
- export-workbook-to-html-after-defining-a-print-area-and-disabling-css-generation.cs
- set-htmlsaveoptionssimilarborderstyle-to-true-to-provide-fallback-borders-for-unsupported-browsers-in-html.cs
- use-htmlsaveoptionssimilarborderstyle-to-generate-fallback-border-styles-and-test-rendering-in-legacy-browsers.cs
- test-fallback-border-rendering-in-firefox-by-exporting-with-similarborderstyle-enabled-and-inspecting-css.cs
- test-fallback-border-rendering-in-safari-by-exporting-with-similarborderstyle-enabled-and-reviewing-html-output.cs
- enable-htmlsaveoptionsisiecompatible-before-saving-to-mhtml-to-allow-worksheet-tab-switching-in-internet-explorer.cs
- generate-iecompatible-mhtml-from-a-multisheet-workbook-and-verify-tab-switching-functionality-in-internet-explorer.cs
- test-mhtml-output-in-microsoft-edge-after-enabling-isiecompatible-and-assess-compatibility-with-modern-browsers.cs
- test-mhtml-output-in-google-chrome-with-isiecompatible-false-and-confirm-standard-rendering-without-iespecific-features.cs
- implement-ifilepathprovider-to-customize-html-file-paths-for-each-exported-worksheet-and-maintain-link-integrity.cs
- assign-a-custom-ifilepathprovider-instance-to-htmlsaveoptionsfilepathprovider-to-fix-broken-worksheet-hyperlinks.cs
- implement-ifilepathprovider-that-converts-worksheet-names-to-lowercase-urls-for-consistent-linking.cs
- implement-ifilepathprovider-that-generates-unique-file-names-for-worksheets-with-duplicate-titles.cs
- store-exported-html-files-in-a-temporary-folder-via-ifilepathprovider-and-clean-up-after-processing.cs
- export-workbook-to-html-with-custom-file-naming-pattern-using-ifilepathprovider-that-appends-timestamps.cs
- export-workbook-to-html-while-preserving-original-worksheet-order-by-not-altering-ifilepathprovider-mapping.cs
- use-ifilepathprovider-to-store-html-files-in-a-network-share-path-and-validate-accessibility-via-unc.cs
- use-ifilepathprovider-to-generate-relative-paths-for-html-files-to-support-offline-browsing.cs
- set-htmlsaveoptionsexportworksheetcssseparately-to-true-to-generate-distinct-stylesheet-files-per-worksheet-individually.cs
- export-a-singlesheet-workbook-to-html-with-exportworksheetcssseparately-enabled-creating-a-css-folder.cs
- create-external-css-files-per-worksheet-and-ensure-html-references-point-to-the-correct-stylesheet-locations.cs
- disable-embedded-css-and-manually-link-an-external-stylesheet-after-html-export-for-custom-styling.cs
- use-htmlsaveoptionsexportworksheetcssseparately-together-with-similarborderstyle-to-test-combined-effects-on-rendering-across-multiple-browsers.cs
- use-htmlsaveoptionsexportworksheetcssseparately-true-and-verify-that-each-worksheet-has-its-own-css-file.cs
- use-htmlsaveoptionsexportworksheetcssseparately-false-and-confirm-that-css-is-embedded-within-the-html.cs
- export-workbook-to-html-with-custom-css-folder-path-and-ensure-the-folder-is-created-relative-to-output.cs
- export-workbook-to-html-with-embedded-css-and-compare-file-size-against-external-css-approach.cs
- load-an-html-file-from-disk-and-convert-it-to-pdf-using-default-options.cs
- load-an-html-string-in-memory-and-generate-a-pdf-with-embedded-fonts.cs
- convert-html-with-external-css-files-to-pdf-while-preserving-original-stylesheet-rules.cs
- convert-html-containing-base64-images-to-pdf-ensuring-images-retain-original-resolution.cs
- generate-pdf-from-html-and-set-page-size-to-a4-landscape-orientation.cs
- convert-html-to-pdf-applying-custom-margins-of-one-centimeter-on-all-sides.cs
- convert-html-to-pdf-while-preserving-hyperlinks-and-making-them-clickable.cs
- convert-html-to-pdf-with-pdfa1b-compliance-for-longterm-archival-storage.cs
- convert-html-to-pdf-and-encrypt-the-output-file-with-a-user-password.cs
- convert-html-to-pdf-and-add-a-visible-watermark-text-across-each-page.cs
- convert-html-to-pdf-and-embed-document-metadata-such-as-title-and-author.cs
- convert-html-to-pdf-with-high-image-compression-level-to-reduce-file-size.cs
- convert-html-to-pdf-using-a-dpi-setting-of-300-for-highquality-graphics.cs
- convert-html-to-pdf-and-enable-vector-graphics-rendering-for-scalable-charts.cs
- convert-html-to-pdf-while-preserving-righttoleft-text-direction-for-arabic-content.cs
- convert-html-to-pdf-and-include-accessibility-tags-for-screen-reader-compatibility.cs
- convert-html-to-pdf-with-custom-font-mapping-to-replace-missing-fonts.cs
- convert-html-to-pdf-and-add-a-header-with-the-source-file-name-on-each-page.cs
- convert-html-to-pdf-and-add-a-footer-displaying-page-numbers-in-roman-numerals.cs
- convert-html-to-pdf-and-set-pdf-version-to-17-for-advanced-features.cs
- convert-html-to-pdf-asynchronously-and-report-conversion-progress-to-callers-via-events.cs
- convert-html-to-pdf-in-a-multithreaded-batch-processing-fifty-files-simultaneously.cs
- convert-html-to-pdf-and-apply-a-digital-signature-using-a-provided-certificate.cs
- convert-html-to-pdf-and-compress-the-final-document-using-object-stream-compression.cs
- convert-html-to-pdf-while-preserving-gradient-fills-in-css-background-images.cs
- convert-html-to-pdf-and-retain-embedded-video-placeholders-as-static-images.cs
- convert-html-to-pdf-and-ensure-table-borders-render-with-exact-pixel-widths.cs
- convert-html-to-pdf-and-preserve-css-pseudoelements-like-before-and-after.cs
- convert-html-to-pdf-and-maintain-original-line-spacing-and-paragraph-indentation.cs
- convert-html-to-pdf-and-embed-a-custom-icc-color-profile-for-accurate-reproduction.cs
- load-html-from-a-url-and-convert-it-to-an-excel-workbook-using-a-custom-stream-provider.cs
- convert-html-file-to-excel-while-preserving-cell-background-colors-defined-by-css-styles.cs
- convert-html-tables-with-colspan-and-rowspan-attributes-to-merged-cells-in-the-worksheet.cs
- convert-html-with-inline-styles-to-excel-and-map-css-font-sizes-to-row-heights.cs
- convert-html-with-embedded-images-to-excel-inserting-each-image-into-the-corresponding-cell.cs
- convert-html-to-excel-and-retain-hyperlinks-making-them-clickable-within-the-workbook.cs
- convert-html-to-excel-while-preserving-numeric-formats-such-as-currency-and-percentages.cs
- convert-html-to-excel-and-apply-date-format-detection-based-on-locale-settings.cs
- convert-html-to-excel-and-map-css-text-alignment-to-excel-cell-alignment-properties.cs
- convert-html-to-excel-using-a-stream-provider-that-writes-output-directly-to-cloud-storage.cs
- convert-html-to-excel-and-encrypt-the-resulting-workbook-with-a-password-for-protection.cs
- convert-html-to-excel-and-preserve-conditional-formatting-rules-defined-in-css-classes.cs
- convert-html-to-excel-and-embed-custom-document-properties-extracted-from-html-meta-tags.cs
- convert-html-to-excel-and-retain-embedded-svg-graphics-as-scalable-vector-shapes.cs
- convert-html-to-excel-and-apply-custom-number-formats-for-scientific-notation-values.cs
- convert-html-to-excel-and-preserve-cell-comments-as-html-tooltip-attributes.cs
- convert-html-to-excel-and-handle-nested-tables-by-creating-separate-worksheets-for-each-level.cs
- convert-html-to-excel-and-maintain-original-html-page-breaks-as-worksheet-page-breaks.cs
- convert-html-to-excel-and-set-workbook-culture-to-french-for-proper-decimal-separators.cs
- convert-html-to-excel-using-a-memory-stream-to-process-large-files-without-temporary-files.cs
- convert-html-to-excel-asynchronously-and-provide-a-callback-when-conversion-completes.cs
- convert-html-to-excel-in-a-batch-job-processing-a-directory-of-files-into-separate-workbooks.cs
- convert-html-to-excel-and-apply-a-custom-theme-based-on-css-variables-defined-in-the-html.cs
- convert-html-to-excel-and-preserve-hidden-rows-and-columns-indicated-by-css-displaynone.cs
- convert-html-to-excel-and-map-css-border-styles-to-excel-cell-border-line-styles.cs
- convert-html-to-excel-and-retain-background-images-by-inserting-them-as-worksheet-background.cs
- convert-html-to-excel-and-set-workbook-calculation-mode-to-automatic-for-formula-evaluation.cs
- load-an-excel-workbook-from-a-file-and-export-it-to-html-using-a-custom-stream-provider.cs
- export-excel-to-html-while-preserving-cell-formatting-such-as-font-styles-and-colors.cs
- export-excel-to-html-and-generate-a-separate-css-file-for-styling-instead-of-inline-styles.cs
- export-excel-to-html-and-embed-images-as-base64-data-uris-within-the-html-output.cs
- export-excel-to-html-and-retain-merged-cells-using-appropriate-colspan-and-rowspan-attributes.cs
- export-excel-to-html-and-preserve-worksheet-gridlines-as-css-border-definitions.cs
- export-excel-to-html-and-include-a-custom-html-header-containing-the-workbook-title.cs
- export-excel-to-html-and-add-a-footer-with-page-numbers-generated-by-javascript.cs
- export-excel-to-html-and-set-the-output-encoding-to-utf8-for-international-character-support.cs
- export-excel-to-html-and-apply-responsive-design-techniques-for-optimal-mobile-device-viewing.cs
- export-excel-to-html-and-embed-hyperlinks-so-they-remain-functional-in-the-browser.cs
- export-excel-to-html-and-include-custom-metadata-tags-extracted-from-workbook-properties.cs
- export-excel-to-html-and-compress-the-output-html-using-gzip-stream-for-faster-transmission.cs
- export-excel-to-html-and-apply-a-custom-stylesheet-that-overrides-default-cell-colors.cs
- export-excel-to-html-and-preserve-cell-comments-as-html-tooltip-attributes.cs
- export-excel-to-html-and-ensure-formulas-are-displayed-as-values-not-as-formula-strings.cs
- export-excel-to-html-and-maintain-original-column-widths-using-css-width-properties.cs
- export-excel-to-html-and-set-page-break-markers-as-html-comments-for-later-processing.cs
- export-excel-to-html-and-include-a-navigation-pane-linking-to-each-worksheet-section.cs
- export-excel-to-html-and-handle-hidden-rows-by-omitting-them-from-the-generated-markup.cs
- export-excel-to-html-and-preserve-background-images-by-linking-to-external-image-files.cs
- export-excel-to-html-and-apply-localespecific-number-formatting-based-on-workbook-culture.cs
- export-excel-to-html-and-generate-a-printable-version-using-css-media-print-rules.cs
- export-excel-to-html-and-embed-custom-javascript-for-interactive-sorting-of-table-columns.cs
- export-excel-to-html-and-ensure-merged-cells-spanning-multiple-rows-render-correctly.cs
- convert-a-spreadsheet-containing-wordart-to-pdf-and-render-gradient-fills-accurately-in-the-output.cs
- convert-a-spreadsheet-with-wordart-to-svg-and-ensure-gradient-definitions-are-preserved-as-vector-data.cs
- render-wordart-gradient-fills-during-pdf-conversion-and-enable-antialiasing-for-smoother-edges.cs
- render-wordart-gradient-fills-during-html-export-and-apply-custom-css-classes-for-gradient-control.cs
- render-wordart-gradient-fills-while-converting-to-pdf-and-embed-the-gradient-as-a-pdf-shading-pattern.cs
- render-wordart-gradient-fills-while-converting-to-html-and-include-fallback-solid-colors-for-older-browsers.cs
- render-wordart-gradient-fills-during-conversion-and-allow-configuration-of-gradient-direction-via-api.cs
- render-wordart-gradient-fills-while-converting-and-support-custom-start-and-end-colors-defined-in-the-workbook.cs
- convert-multiple-spreadsheets-containing-wordart-to-pdf-in-a-batch-ensuring-each-gradient-renders-correctly.cs
- convert-multiple-spreadsheets-with-wordart-to-html-in-a-batch-generating-separate-css-files-for-gradients.cs
- validate-that-gradient-fills-in-wordart-appear-identical-in-pdf-output-compared-to-the-original-spreadsheet.cs
- apply-a-shadow-effect-to-wordart-during-pdf-conversion-while-preserving-the-underlying-gradient-fill.cs
- apply-a-reflection-effect-to-wordart-during-html-export-and-ensure-gradient-fill-remains-visible.cs
- customize-gradient-angle-for-wordart-during-pdf-conversion-using-a-specific-rotation-parameter.cs
- customize-gradient-angle-for-wordart-during-html-conversion-and-reflect-the-angle-in-css-gradient-syntax.cs
- convert-a-spreadsheet-with-wordart-to-pdf-and-embed-the-gradient-fill-as-an-iccprofileaware-color-space.cs
- convert-a-spreadsheet-with-wordart-to-html-and-include-inline-svg-definitions-for-complex-gradients.cs
- perform-asynchronous-conversion-of-wordartrich-spreadsheets-to-pdf-reporting-progress-of-gradient-rendering.cs
- perform-asynchronous-conversion-of-wordartrich-spreadsheets-to-html-handling-gradient-generation-on-separate-threads.cs
- convert-a-workbook-with-wordart-to-pdf-and-add-a-security-restriction-preventing-editing-of-gradient-layers.cs
- convert-a-workbook-with-wordart-to-html-and-minify-the-generated-css-defining-gradient-fills.cs
- convert-a-workbook-with-wordart-to-pdf-and-set-output-to-pdfa2b-compliance-while-preserving-gradients.cs
- convert-a-workbook-with-wordart-to-html-and-ensure-output-passes-w3c-validation-for-css-gradients.cs
- render-wordart-gradient-fills-during-pdf-conversion-and-embed-a-fallback-raster-image-for-unsupported-viewers.cs
- render-wordart-gradient-fills-while-converting-to-html-and-provide-a-javascript-fallback-that-draws-gradients-on-canvas.cs
- convert-a-spreadsheet-containing-multiple-wordart-objects-to-pdf-and-verify-each-gradient-renders-independently.cs
- load-a-workbook-from-an-excel-file-and-export-to-html-using-a-custom-tablecssid.cs
- export-the-workbook-with-gridlines-enabled-while-applying-a-specified-tablecssid-for-table-styling.cs
- include-worksheet-headings-in-the-html-output-and-prefix-table-styles-with-the-chosen-tablecssid.cs
- disable-default-css-generation-and-rely-solely-on-tablecssid-to-style-exported-html-tables.cs
- enable-css-custom-properties-in-htmlsaveoptions-while-using-tablecssid-to-enhance-table-styling-flexibility.cs
- include-cell-comments-in-the-html-export-and-ensure-they-inherit-the-tablecssid-prefixed-styles.cs
- set-the-html-encoding-to-utf-8-and-apply-a-custom-tablecssid-for-consistent-table-class-naming.cs
- batch-process-multiple-workbooks-exporting-them-to-html-using-a-shared-tablecssid-to-maintain-uniform-styling.cs
- batch-export-several-workbooks-to-html-assigning-each-a-distinct-tablecssid-to-differentiate-table-class-prefixes.cs
- link-the-generated-html-to-an-external-stylesheet-that-targets-the-prefixed-tablecssid-classes-for-sitewide-styling.cs
- validate-the-exported-html-using-an-html-validator-to-confirm-correct-tablecssid-prefixes-on-table-elements.cs
- programmatically-change-the-tablecssid-after-initial-export-resave-the-html-and-verify-updated-class-names.cs
- use-workbooksaveashtml-shortcut-with-preconfigured-htmlsaveoptions-that-include-a-custom-tablecssid-for-quick-export.cs
- load-an-html-file-into-a-workbook-then-reexport-it-to-html-with-a-different-tablecssid-applied.cs
- apply-conditional-formatting-in-the-source-workbook-and-verify-its-visual-preservation-in-html-output-using-tablecssid.cs
- enable-gridlines-during-export-to-retain-cell-borders-and-ensure-tablecssid-prefixes-style-those-borders-correctly.cs
- export-a-workbook-containing-hidden-rows-and-columns-confirming-that-tablecssid-prefixes-only-appear-on-visible-tables.cs
- hide-overlaid-content-using-crosshideright-while-exporting-to-html-and-apply-a-custom-tablecssid-for-table-styling.cs
- recognize-selfclosing-tags-in-the-html-output-and-ensure-they-are-correctly-prefixed-with-the-tablecssid-identifier.cs
- autofit-columns-and-rows-after-loading-html-into-a-workbook-then-export-with-tablecssid-to-preserve-layout.cs
- export-only-the-active-worksheet-to-html-while-applying-a-specific-tablecssid-to-style-its-tables.cs
- generate-html-with-a-custom-tablecssid-and-verify-that-no-css-class-name-collisions-occur-with-existing-styles.cs
- create-inline-css-definitions-within-the-exported-html-and-prefix-them-using-tablecssid-for-immediate-preview.cs
- export-a-worksheet-containing-multiple-tables-and-ensure-each-table-receives-the-tablecssid-prefixed-style-class.cs
- export-a-workbook-with-hidden-worksheets-and-confirm-that-tablecssid-prefixed-styles-are-generated-only-for-visible-sheets.cs
- include-a-timestamp-in-the-html-filename-and-retain-the-tablecssid-prefix-to-track-export-versions-over-time.cs
- log-each-export-operation-with-details-such-as-source-workbook-tablecssid-used-and-destination-path-for-audit-purposes.cs
- implement-exception-handling-to-catch-invalid-tablecssid-values-and-provide-descriptive-error-messages-to-developers.cs
- load-a-workbook-enable-widthscalable-and-save-as-html-with-percentage-column-widths.cs
- export-a-workbook-to-html-with-widthscalable-set-to-true-for-embased-column-sizing.cs
- apply-widthscalable-false-and-confirm-column-widths-render-as-fixed-pixel-values.cs
- apply-widthscalable-true-and-confirm-column-widths-adapt-using-em-units.cs
- generate-html-using-htmlcrosstypedefault-to-mimic-excel-overflow-for-long-strings.cs
- produce-html-with-htmlcrosstypefittocell-to-restrict-text-overflow-within-cell-bounds.cs
- create-html-using-htmlcrosstypecross-for-highperformance-export-of-large-workbooks.cs
- export-a-workbook-to-html-with-htmlcrosstypemsexport-to-replicate-excels-native-html-style.cs
- confirm-that-htmlcrosstypefittocell-wraps-text-within-cell-boundaries-when-enabled.cs
- measure-export-time-for-a-large-workbook-using-htmlcrosstypecross-versus-default.cs
- validate-that-htmlcrosstypecross-improves-performance-without-altering-visual-rendering-compared-to-default.cs
- save-html-with-addtooltiptext-enabled-so-full-cell-text-appears-as-hover-tooltip.cs
- generate-html-with-addtooltiptext-disabled-to-improve-rendering-speed-significantly.cs
- verify-tooltip-attributes-appear-only-for-cells-whose-displayed-text-is-truncated.cs
- ensure-cells-with-short-text-do-not-receive-tooltip-attributes-when-addtooltiptext-is-enabled.cs
- combine-widthscalable-htmlcrosstypefittocell-and-addtooltiptext-to-produce-responsive-html-with-tooltips.cs
- use-htmlcrosstypefittocell-together-with-addtooltiptext-to-provide-full-content-on-hover-without-overflow.cs
- use-htmlcrosstypedefault-with-addtooltiptext-enabled-to-show-full-cell-content-while-preserving-overflow-behavior.cs
- export-only-visible-worksheets-by-setting-exporthiddenworksheet-to-false-before-saving.cs
- include-hidden-worksheets-by-leaving-exporthiddenworksheet-set-to-true-during-html-export.cs
- export-html-with-exporthiddenworksheet-false-while-using-widthscalable-true-for-responsive-columns.cs
- export-html-with-exporthiddenworksheet-true-and-htmlcrosstypecross-for-performance-on-large-workbooks.cs
- verify-that-hidden-worksheets-are-absent-from-the-generated-html-when-exporthiddenworksheet-is-false.cs
- ensure-that-setting-exporthiddenworksheet-to-true-includes-hidden-worksheets-in-the-html-output.cs
- generate-html-with-widthscalable-enabled-and-addtooltiptext-enabled-to-create-scalable-columns-with-hover-tooltips.cs
- compare-html-file-size-when-widthscalable-is-true-versus-false-for-the-same-workbook.cs
- test-that-enabling-widthscalable-does-not-affect-nonhtml-export-formats.cs
- create-a-utility-method-that-loads-a-workbook-sets-addtooltiptext-true-and-returns-html-string.cs
- batch-convert-a-folder-of-excel-files-to-html-using-default-htmlsaveoptions-and-log-conversion-errors.cs
- convert-multiple-workbooks-in-a-loop-applying-different-htmlcrosstype-values-based-on-file-size.cs
- batch-process-workbooks-applying-exporthiddenworksheet-false-only-to-those-containing-confidential-worksheets.cs
- verify-that-exporthiddenworksheet-false-correctly-hides-hidden-sheets-while-preserving-visible-content.cs
- ensure-that-combining-exporthiddenworksheet-false-with-htmlcrosstypefittocell-limits-overflow-and-hides-hidden-worksheets.cs
- validate-that-htmlcrosstypecross-improves-performance-for-workbooks-containing-thousands-of-rows.cs
- convert-an-excel-workbook-to-html-using-default-htmlsaveoptions.cs
- convert-an-excel-workbook-to-html-with-presentationpreferencebestfit-for-improved-layout.cs
- convert-an-excel-workbook-to-html-with-presentationpreferenceautofit-for-compact-column-widths.cs
- convert-an-excel-workbook-to-html-with-presentationpreferencebestfit-while-exporting-cell-comments.cs
- convert-an-excel-workbook-to-html-with-presentationpreferencebestfit-and-render-gridlines.cs
- convert-an-excel-workbook-to-html-with-presentationpreferencebestfit-and-export-hidden-worksheets.cs
- convert-an-excel-workbook-to-html-with-presentationpreferencebestfit-and-enable-css-custom-properties-for-image-deduplication.cs
- convert-an-excel-workbook-to-html-with-presentationpreferencebestfit-while-disabling-all-css-generation.cs
- convert-an-excel-workbook-to-html-with-presentationpreferencebestfit-and-set-hyperlink-target-to-blank.cs
- convert-an-excel-workbook-to-html-with-presentationpreferencebestfit-and-set-hyperlink-target-to-parent.cs
- convert-an-excel-workbook-to-html-with-default-options-and-set-hyperlink-target-to-blank.cs
- convert-an-excel-workbook-to-html-with-default-options-and-set-hyperlink-target-to-parent.cs
- convert-an-excel-workbook-to-html-with-default-options-and-enable-css-custom-properties-for-image-reuse.cs
- convert-an-excel-workbook-to-html-with-default-options-and-disable-css-generation-for-minimal-output.cs
- convert-an-excel-workbook-to-html-with-default-options-while-exporting-cell-comments.cs
- convert-an-excel-workbook-to-html-with-default-options-while-exporting-conditional-formatting.cs
- convert-an-excel-workbook-to-html-with-default-options-while-exporting-gridlines.cs
- convert-an-excel-workbook-to-html-with-default-options-while-exporting-hidden-worksheets.cs
- batch-convert-all-excel-files-in-a-folder-to-html-using-default-htmlsaveoptions.cs
- batch-convert-excel-workbooks-to-html-with-presentationpreferencebestfit-for-each-file.cs
- batch-convert-excel-files-to-html-enabling-css-custom-properties-to-deduplicate-images.cs
- batch-convert-excel-workbooks-to-html-while-disabling-css-generation-for-lightweight-files.cs
- batch-convert-excel-files-to-html-with-exportcomments-enabled-to-include-all-cell-notes.cs
- batch-convert-excel-workbooks-to-html-with-exportconditionalformatting-enabled-to-retain-visual-rules.cs
- batch-convert-excel-files-to-html-with-exportgridlines-enabled-to-display-spreadsheet-grid-structure.cs
- batch-convert-excel-workbooks-to-html-with-exporthiddenworksheet-enabled-to-include-hidden-sheet-data.cs
- validate-generated-html-contains-link-target-attribute-_blank-after-setting-linktargettype-to-blank.cs
- validate-generated-html-contains-link-target-attribute-_parent-after-setting-linktargettype-to-parent.cs
- verify-css-custom-properties-reduce-duplicate-base64-image-strings-compared-to-default-output.cs
- compare-html-file-size-when-enablecsscustomproperties-is-true-versus-false-for-same-workbook.cs
- compare-html-file-size-when-disablecss-is-true-versus-false-to-assess-markup-reduction.cs
- ensure-cell-comments-appear-in-html-output-when-exportcomments-option-is-enabled.cs
- ensure-conditional-formatting-colors-are-retained-in-html-when-exportconditionalformatting-is-true.cs
- ensure-gridlines-are-visible-in-html-output-when-exportgridlines-option-is-enabled.cs
- ensure-hidden-worksheet-content-appears-in-html-when-exporthiddenworksheet-option-is-true.cs
- measure-rendering-performance-of-html-with-css-disabled-versus-enabled-using-a-stopwatch-timer.cs
- generate-html-with-a-custom-page-title-by-setting-htmlsaveoptionshtmltitle-property.cs
- generate-html-using-utf-8-encoding-by-assigning-htmlsaveoptionsencoding-to-encodingutf8.cs
- save-html-output-to-a-memorystream-instead-of-a-physical-file-for-inmemory-processing.cs
- write-html-bytes-from-a-memorystream-directly-to-an-aspnet-response-for-immediate-download.cs
- append-a-custom-css-class-to-the-html-body-element-after-conversion-for-additional-styling.cs
- replace-all-hyperlink-urls-in-generated-html-with-absolute-urls-using-string-replacement-logic.cs
- set-exporthiddenworksheet-false-and-verify-hidden-sheets-are-omitted-from-resulting-html.cs
- set-exporthiddenworksheet-true-and-verify-hidden-sheets-are-included-in-generated-html.cs
- create-a-unit-test-asserting-a-specific-cell-value-appears-in-the-html-output.cs
- create-a-unit-test-asserting-no-style-tags-exist-when-disablecss-option-is-true.cs
- create-a-unit-test-asserting-css-custom-property-definitions-appear-within-a-root-selector-when-enabled.cs
- generate-html-for-a-workbook-containing-merged-cells-and-verify-merged-layout-is-preserved.cs
- generate-html-for-a-workbook-with-frozen-panes-and-verify-pane-positions-are-reflected.cs
- generate-html-for-a-workbook-with-data-validation-dropdowns-and-verify-they-appear-as-select-elements.cs
- generate-html-for-a-workbook-with-formulas-and-verify-calculated-results-are-displayed-correctly.cs
- generate-html-for-a-workbook-with-hidden-rows-and-verify-they-are-omitted-unless-exporthiddenworksheet-is-true.cs
- convert-a-workbook-to-html-then-reload-it-using-loadoptions-to-verify-roundtrip-fidelity.cs
- compare-visual-fidelity-of-html-generated-with-presentationpreferencebestfit-versus-default-layout.cs
- measure-memory-usage-when-converting-a-large-workbook-to-html-with-css-enabled.cs
- measure-memory-usage-when-converting-the-same-large-workbook-to-html-with-css-disabled.cs
- set-exportcomments-false-and-verify-comment-markers-are-absent-from-generated-html.cs
- set-exportconditionalformatting-false-and-verify-conditional-style-attributes-are-missing-in-html.cs
- set-exportgridlines-false-and-verify-gridline-elements-are-not-rendered-in-html-output.cs
- set-exporthiddenworksheet-false-and-verify-hidden-worksheets-are-excluded-from-html-conversion.cs
- combine-exportcomments-false-exportconditionalformatting-false-and-exportgridlines-false-to-generate-minimal-html.cs
- generate-html-with-embedded-images-as-separate-files-by-setting-exportimagesasbase64-false.cs
- generate-html-with-embedded-images-as-base64-strings-by-setting-exportimagesasbase64-true.cs
- verify-base64-image-strings-are-deduplicated-when-enablecsscustomproperties-is-true-during-conversion.cs
- verify-duplicate-image-files-are-not-created-when-exportimagesasbase64-is-false-during-conversion.cs
- create-a-script-that-extracts-excel-files-from-a-zip-archive-and-converts-each-to-html.cs
- implement-error-handling-for-missing-worksheets-when-exporthiddenworksheet-option-is-enabled.cs
- implement-logging-of-conversion-duration-for-each-workbook-during-batch-html-processing.cs
- set-exportactiveworksheetonly-true-to-generate-html-containing-only-the-currently-active-worksheet.cs
- set-exportallworksheets-true-to-generate-a-single-html-file-that-includes-all-workbook-worksheets.cs
- set-worksheetindex-in-htmlsaveoptions-to-export-a-specific-worksheet-by-its-zerobased-position.cs
- set-sheetname-in-htmlsaveoptions-to-export-a-specific-worksheet-identified-by-its-name.cs
- generate-html-and-then-compress-it-using-gzipstream-for-efficient-web-transmission.cs
- generate-html-and-embed-it-in-an-email-body-using-systemnetmail-for-newsletter-distribution.cs
- generate-html-and-compare-its-dom-structure-against-an-expected-xml-document-to-validate-layout.cs
- set-exportchartimageformat-to-png-and-verify-chart-images-in-html-are-png-files.cs
- set-exportchartimageformat-to-jpeg-and-verify-chart-images-in-html-are-jpeg-files.cs
- generate-html-and-then-modify-the-output-to-add-a-custom-css-class-to-header-rows.cs
- generate-html-and-ensure-numeric-values-retain-their-original-formatting-as-displayed-in-excel.cs
- generate-html-and-ensure-date-values-are-formatted-according-to-the-workbooks-locale-settings.cs
- generate-html-and-ensure-text-wrapping-is-preserved-in-cells-where-wrap-text-is-enabled.cs
- generate-html-and-ensure-font-styles-such-as-bold-and-italic-are-retained-with-css-enabled.cs
- generate-html-and-ensure-hyperlinks-retain-their-original-url-targets-after-conversion.cs
- generate-html-and-ensure-table-headers-include-column-letters-when-exporting-column-headers.cs
- generate-html-and-ensure-row-numbers-are-displayed-when-exportrowheaders-option-is-enabled.cs
- generate-html-and-ensure-column-letters-are-displayed-when-exportcolumnheaders-option-is-enabled.cs
- create-a-console-application-that-accepts-an-excel-file-path-and-outputs-html-using-bestfit-layout.cs
- create-a-windows-service-that-monitors-a-directory-and-converts-new-excel-files-to-html-with-custom-options.cs
- write-a-powershell-script-that-invokes-the-net-library-to-batch-convert-excel-files-to-html.cs
- integrate-html-conversion-into-an-aspnet-mvc-controller-action-that-returns-a-fileresult-containing-the-html.cs
