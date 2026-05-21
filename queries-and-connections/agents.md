# Queries and Connections Examples

This folder contains **Aspose.Cells for .NET** code examples related to:

Queries and Connections


## Purpose

These examples demonstrate common **Aspose.Cells APIs** used when working with:

- Workbooks
- Worksheets
- Cells
- Formulas
- Charts
- Data operations


## Example Files

Each `.cs` file demonstrates a specific task related to **Queries and Connections**.

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
- load-a-workbook-from-a-file-path-and-obtain-its-dataconnections-collection.cs
- iterate-through-dataconnections-to-identify-sql-type-connections-and-cast-each-to-dbconnection.cs
- retrieve-commandtext-commandtype-and-connectioninfo-from-each-dbconnection-for-inspection.cs
- update-username-and-password-properties-of-a-specific-dbconnection-with-new-credentials.cs
- change-description-property-of-a-dbconnection-to-reflect-its-new-purpose-after-migration.cs
- modify-commandtext-of-a-dbconnection-to-include-a-filter-clause-limiting-returned-rows.cs
- set-refreshonload-flag-of-a-webquery-connection-to-true-for-automatic-data-update.cs
- enable-backgroundrefresh-on-a-webquery-connection-to-avoid-blocking-ui-during-retrieval.cs
- disable-backgroundrefresh-on-a-sql-dbconnection-to-enforce-sequential-query-execution.cs
- refresh-all-external-data-connections-sequentially-to-ensure-data-consistency-across-the-workbook.cs
- retrieve-the-first-worksheet-access-its-first-querytable-and-read-preserveformatting-flag.cs
- enable-preserveformatting-on-a-querytable-to-maintain-cell-styles-after-data-refresh.cs
- disable-preserveformatting-on-a-querytable-to-allow-default-formatting-during-subsequent-refreshes.cs
- iterate-through-all-querytables-in-all-worksheets-and-set-preserveformatting-to-true.cs
- obtain-resultrange-address-of-a-querytable-and-log-it-for-downstream-processing.cs
- rename-an-existing-dbconnection-to-reflect-a-new-database-server-after-migration.cs
- validate-that-all-external-connections-have-nonempty-credentials-before-saving-the-workbook.cs
- save-the-modified-workbook-to-a-new-file-path-preserving-original-version-metadata.cs
- write-a-utility-that-lists-all-worksheets-containing-query-tables-and-outputs-their-names.cs
- measure-and-log-the-time-taken-to-refresh-each-dbconnection-for-performance-analysis.cs
- load-an-xls-workbook-using-the-workbook-class-and-access-its-datamashup-property.cs
- extract-odata-connection-details-from-the-datamashup-object-and-log-the-service-endpoint-url.cs
- output-each-powerquery-formula-name-to-the-console-for-verification-purposes.cs
- retrieve-the-first-external-dbconnection-from-a-workbook-and-read-its-current-name-property.cs
- rename-the-retrieved-dbconnectionname-to-a-descriptive-identifier-such-as-salesdataconnection.cs
- change-the-path-property-of-an-external-link-in-an-xls-workbook-to-a-network-shared-folder.cs
- detect-hidden-external-links-within-the-workbook-using-the-appropriate-api-and-list-their-source-paths.cs
- generate-a-plaintext-report-summarizing-hidden-external-link-paths-for-further-analysis.cs
- save-the-modified-workbook-as-an-xlsb-file-while-preserving-all-external-connection-settings.cs
- validate-that-updating-a-powerqueryformulaitemvalue-correctly-changes-the-underlying-query-definition.cs
- load-multiple-xls-files-from-a-directory-update-each-dbconnectionname-and-save-changes-in-place.cs
- export-a-list-of-all-external-connection-names-from-a-workbook-to-a-plain-text-file.cs
- compare-odata-connection-metadata-before-and-after-modification-to-ensure-version-consistency.cs
- use-workbookdatamashup-to-extract-odata-service-urls-and-store-them-in-a-json-configuration-file.cs
- programmatically-remove-a-hidden-external-link-from-the-workbook-and-verify-its-absence.cs
- set-the-path-property-of-an-external-link-to-a-relative-path-and-test-workbook-portability.cs
- serialize-the-powerqueryformulacollection-to-xml-for-external-auditing-compliance-purposes.cs
- load-an-xlsb-workbook-change-the-dbconnectionname-and-log-the-modification-timestamp.cs
- detect-and-list-external-connections-of-type-webquery-across-a-batch-of-workbooks.cs
- update-the-source-file-location-of-a-power-query-data-source-to-a-cloud-storage-url.cs
- verify-that-hidden-external-links-remain-hidden-after-encrypting-the-workbook-with-a-password.cs
- add-error-handling-to-capture-exceptions-when-an-external-connection-path-is-invalid.cs
- generate-a-csv-file-containing-workbook-name-connection-type-and-connection-name-for-all-files-in-a-folder.cs
- replace-all-occurrences-of-a-deprecated-database-name-within-dbconnectionname-properties-across-workbooks.cs
- change-the-absolute-path-of-an-external-link-data-source-file-programmatically-for-a-workbook.cs
- validate-that-the-new-external-link-path-points-to-an-existing-file-before-applying-changes.cs
- update-all-external-links-in-the-workbook-to-use-relative-paths-for-better-portability.cs
- log-the-original-and-updated-external-link-paths-for-audit-purposes.cs
- handle-errors-when-the-external-link-file-is-missing-or-inaccessible-during-path-update.cs
- create-a-backup-of-the-workbook-before-modifying-external-link-paths.cs
- add-a-custom-ribbon-tab-named-data-tools-using-ribbon-xml-definition.cs
