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
