# Globalization and Localization Examples

This folder contains **Aspose.Cells for .NET** code examples related to:

Globalization and Localization


## Purpose

These examples demonstrate common **Aspose.Cells APIs** used when working with:

- Workbooks
- Worksheets
- Cells
- Formulas
- Charts
- Data operations


## Example Files

Each `.cs` file demonstrates a specific task related to **Globalization and Localization**.

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
- create-a-custom-globalizationsettings-class-overriding-getlocalfunctionname-for-target-language-functions.cs
- override-geterrorstring-in-the-custom-class-to-provide-localized-error-messages-for-excel-errors.cs
- override-getbooleanstring-to-return-localized-truefalse-strings-for-the-selected-locale.cs
- assign-the-custom-globalizationsettings-instance-to-workbooksettingsglobalizationsettings-before-loading-any-worksheets.cs
- load-the-excel-workbook-using-workbookload-after-configuring-the-custom-globalization-settings.cs
- set-cell-formulas-with-cellformulalocal-to-apply-localized-function-names-during-workbook-processing.cs
- verify-that-localized-function-names-are-correctly-recognized-by-excel-when-the-workbook-is-opened.cs
- save-the-localized-workbook-as-xlsx-preserving-original-formatting-comments-and-cell-styles.cs
- generate-a-report-listing-processed-workbooks-applied-locales-and-any-localization-errors-encountered.cs
- implement-fallback-to-english-function-names-when-a-requested-locale-lacks-a-defined-mapping.cs
- log-each-overridden-method-call-to-a-debug-file-for-troubleshooting-localization-behavior-at-runtime.cs
- create-a-batch-process-that-applies-the-custom-globalization-settings-to-all-workbooks-in-a-folder.cs
- validate-that-boolean-values-display-localized-truefalse-strings-in-cells-containing-logical-formulas.cs
- write-unit-tests-asserting-getlocalfunctionname-returns-expected-localized-equivalents-for-common-functions.cs
- write-unit-tests-for-geterrorstring-covering-standard-excel-error-codes-across-multiple-locales.cs
- write-unit-tests-for-getbooleanstring-covering-true-false-and-null-values-in-different-locales.cs
- ensure-that-cell-comments-retain-their-original-language-while-function-names-are-localized-according-to-settings.cs
- use-workbooksettings-to-enable-automatic-recalculation-after-applying-localized-formulas-to-ensure-correct-results.cs
- configure-workbook-to-use-a-specific-cultureinfo-object-for-date-and-number-formats-during-localization.cs
- test-that-excels-autofilter-works-correctly-with-localized-column-headers-after-applying-globalization.cs
- ensure-that-pivot-tables-reflect-localized-subtotal-labels-when-the-workbook-is-opened-in-the-target-language.cs
- verify-that-chart-titles-and-axis-labels-display-localized-text-when-source-cells-contain-localized-strings.cs
- implement-a-method-to-reset-globalizationsettings-to-default-english-behavior-for-specific-worksheets.cs
- create-a-configuration-file-mapping-locale-identifiers-to-corresponding-custom-globalizationsettings-classes.cs
- load-the-configuration-at-runtime-and-instantiate-the-appropriate-localization-class-based-on-user-input.cs
- develop-a-console-application-that-prompts-users-to-select-a-target-language-and-applies-localization-to-a-workbook.cs
- implement-a-feature-that-switches-localization-at-runtime-based-on-user-selection-without-reloading-the-workbook.cs
- measure-performance-differences-when-loading-workbooks-with-and-without-custom-globalizationsettings-applied.cs
- log-performance-metrics-for-each-localization-step-during-batch-processing-to-identify-bottlenecks.cs
- create-a-diagnostic-tool-that-compares-original-english-formulas-with-localized-versions-for-accuracy-verification.cs
- provide-documentation-examples-showing-how-to-toggle-between-english-and-localized-function-names-using-formulalocal.cs
- load-an-xlsx-workbook-using-loadoptions-with-cultureinfo-set-to-french-preserving-thread-culture.cs
- create-a-subclass-of-globalizationsettings-that-overrides-gettotalname-to-provide-a-localized-subtotal-label.cs
- create-a-subclass-of-globalizationsettings-overriding-getgrandtotalname-to-supply-a-culturespecific-grand-total-label.cs
- develop-a-chartglobalizationsettings-derivative-that-overrides-getothername-for-a-localized-piechart-other-label.cs
- assign-the-custom-globalizationsettings-to-the-workbook-before-adding-subtotals-to-ensure-localized-labels.cs
- add-subtotal-rows-to-the-worksheet-after-assigning-custom-globalizationsettings-verifying-localized-total-labels.cs
- convert-gregorian-date-cells-to-japanese-calendar-dates-with-cellshelper-preserving-era-information-for-each-cell.cs
- save-the-workbook-containing-japanese-era-dates-as-pdf-confirming-era-symbols-appear-correctly-in-the-output.cs
- batch-process-a-folder-of-xlsx-files-loading-each-with-spanish-cultureinfo-and-exporting-charts-to-png-images.cs
- implement-error-handling-for-loadoptionscultureinfo-when-an-unsupported-locale-identifier-is-supplied-during-workbook-loading.cs
- compare-pdf-output-of-a-workbook-loaded-with-invariant-culture-versus-french-culture-to-assess-number-format-differences.cs
- programmatically-set-worksheet-cell-styles-to-display-dates-in-japanese-era-format-after-conversion-then-export-to-pdf.cs
- measure-performance-impact-of-applying-custom-globalization-settings-versus-default-settings-when-generating-large-pivot-tables.cs
- create-an-application-that-loads-a-workbook-with-italian-cultureinfo-adds-subtotals-and-saves-the-result-as-pdf.cs
- implement-a-method-that-switches-globalizationsettings-at-runtime-based-on-userselected-language-before-chart-creation.cs
- generate-a-report-listing-all-cells-converted-to-japanese-dates-including-original-gregorian-values-for-reference.cs
- apply-custom-number-format-strings-to-percentage-cells-after-loading-the-workbook-with-brazilian-portuguese-cultureinfo.cs
- write-a-script-that-extracts-chart-images-applies-a-localized-other-label-and-saves-them-as-pdf-pages.cs
- demonstrate-preserving-the-original-thread-culture-while-loading-a-workbook-with-french-cultureinfo-using-loadoptions.cs
- create-a-utility-that-converts-date-columns-in-multiple-xls-files-to-japanese-era-format-and-outputs-pdfs.cs
- test-that-overriding-getgrandtotalname-does-not-affect-subtotal-labels-when-both-custom-settings-are-applied-simultaneously.cs
- implement-logging-to-capture-which-culturespecific-label-methods-are-called-during-chart-rendering-for-debugging.cs
- design-a-workflow-that-loads-a-workbook-with-arabic-cultureinfo-adds-righttoleft-subtotals-and-saves-as-pdf.cs
