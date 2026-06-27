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
- add-a-custom-http-header-to-a-webquery-connection-for-required-authentication-token.cs
- set-refreshonload-flag-of-a-webquery-connection-to-true-for-automatic-data-update.cs
- disable-backgroundrefresh-on-a-sql-dbconnection-to-enforce-sequential-query-execution.cs
- refresh-all-external-data-connections-sequentially-to-ensure-data-consistency-across-the-workbook.cs
- retrieve-the-first-worksheet-access-its-first-querytable-and-read-preserveformatting-flag.cs
- iterate-through-all-querytables-in-all-worksheets-and-set-preserveformatting-to-true.cs
- obtain-resultrange-address-of-a-querytable-and-log-it-for-downstream-processing.cs
- remove-a-specific-external-connection-from-the-workbook-based-on-its-description-value.cs
- rename-an-existing-dbconnection-to-reflect-a-new-database-server-after-migration.cs
- validate-that-all-external-connections-have-nonempty-credentials-before-saving-the-workbook.cs
- save-the-modified-workbook-to-a-new-file-path-preserving-original-version-metadata.cs
- write-a-utility-that-lists-all-worksheets-containing-query-tables-and-outputs-their-names.cs
- create-a-reusable-method-that-returns-a-dictionary-of-connection-names-and-their-commandtexts.cs
- load-an-xls-workbook-using-the-workbook-class-and-access-its-datamashup-property.cs
- extract-odata-connection-details-from-the-datamashup-object-and-log-the-service-endpoint-url.cs
- load-an-xlsb-workbook-and-enumerate-all-powerquery-formulas-via-the-powerqueryformulacollection.cs
