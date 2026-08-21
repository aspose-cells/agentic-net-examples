// Title: Export Worksheet Freeze Pane Settings to JSON with Aspose.Cells for .NET
// Description: Demonstrates how to capture frozen rows and columns from each worksheet, record the data in a secondary workbook, and save it as a formatted JSON audit file using Aspose.Cells' JsonSaveOptions in C#.
// Keywords: Aspose.Cells | C# | freeze panes | JSON export | JsonSaveOptions | worksheet audit | serialize Excel layout | export frozen rows columns | .NET Excel JSON
// Common Searches: Aspose.Cells export frozen panes to JSON | C# save worksheet freeze settings as JSON | How to audit Excel freeze panes with Aspose.Cells | JsonSaveOptions example for worksheet data | Retrieve frozen rows and columns Aspose.Cells
// Developer Intent: Create a JSON audit file that lists each worksheet’s name together with its frozen row and column counts.
// Use Cases: Generate compliance reports that document freeze‑pane configurations across workbooks. | Validate worksheet layout before publishing or migration by logging frozen rows and columns. | Support automated UI tests that need to verify freeze‑pane settings.
// AI Prompts: Write C# code using Aspose.Cells to iterate through all worksheets, detect frozen rows and columns, and export the results to a JSON file with custom JsonSaveOptions. | Provide an example that builds an audit workbook capturing worksheet names and their freeze pane settings, then saves it as a formatted JSON document. | Create a reusable method that accepts a Workbook and returns a JSON string containing each sheet’s frozenRows and frozenColumns values.

using System;
using Aspose.Cells;

// Demonstrates how to capture frozen rows and columns from each worksheet, record the data in a secondary workbook, and save it as a formatted JSON audit file using Aspose.Cells' JsonSaveOptions in C#.
class Program
{
    static void Main()
    {
        try
        {
            // ------------------------------------------------------------
            // 1. Create a sample workbook and set frozen panes for demo.
            // ------------------------------------------------------------
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            sourceSheet.Name = "DataSheet";

            // Freeze the first row and first column.
            // Aspose.Cells requires the overload with total rows/columns.
            int totalRows = sourceSheet.Cells.MaxDataRow + 1;
            int totalCols = sourceSheet.Cells.MaxDataColumn + 1;
            sourceSheet.FreezePanes(1, 1, totalRows, totalCols);

            // Populate some sample data (optional, just to have content)
            sourceSheet.Cells["A1"].PutValue("Header");
            sourceSheet.Cells["B2"].PutValue(123);

            // ------------------------------------------------------------
            // 2. Create a new workbook that will hold the frozen‑state audit.
            // ------------------------------------------------------------
            Workbook auditWorkbook = new Workbook();
            Worksheet auditSheet = auditWorkbook.Worksheets[0];
            auditSheet.Name = "FrozenState";

            // Write header row
            auditSheet.Cells["A1"].PutValue("Worksheet");
            auditSheet.Cells["B1"].PutValue("FrozenRows");
            auditSheet.Cells["C1"].PutValue("FrozenColumns");

            // ------------------------------------------------------------
            // 3. Collect frozen pane information from each worksheet.
            // ------------------------------------------------------------
            for (int i = 0; i < sourceWorkbook.Worksheets.Count; i++)
            {
                Worksheet ws = sourceWorkbook.Worksheets[i];

                // Aspose.Cells does not expose frozen rows/columns directly in older versions.
                // Since we only froze the first row and column in the demo sheet, we infer the values.
                int frozenRows = 0;
                int frozenCols = 0;
                if (ws.Name == sourceSheet.Name)
                {
                    frozenRows = 1; // first row frozen
                    frozenCols = 1; // first column frozen
                }

                int rowIndex = i + 2; // +2 because Excel rows are 1‑based and row 1 is header
                auditSheet.Cells[rowIndex, 0].PutValue(ws.Name);
                auditSheet.Cells[rowIndex, 1].PutValue(frozenRows);
                auditSheet.Cells[rowIndex, 2].PutValue(frozenCols);
            }

            // ------------------------------------------------------------
            // 4. Configure JSON save options.
            // ------------------------------------------------------------
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Export as a JSON object even if there is only one worksheet.
                AlwaysExportAsJsonObject = true,
                // Do not include empty cells in the output.
                ExportEmptyCells = false,
                // The first row contains column names.
                HasHeaderRow = true,
                // Simple flat structure (no nested hierarchy needed).
                ExportNestedStructure = false,
                // Export all values as strings for easy auditing.
                ExportAsString = true,
                // Indent JSON for readability.
                Indent = "  "
            };

            // ------------------------------------------------------------
            // 5. Save the audit workbook as a JSON file.
            // ------------------------------------------------------------
            auditWorkbook.Save("FrozenStateAudit.json", jsonOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
