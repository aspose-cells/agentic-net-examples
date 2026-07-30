// Title: Aspose.Cells for .NET – Serialize Worksheet Freeze‑Pane Settings to JSON (C#)
// Description: C# example that creates a workbook with two worksheets, applies FreezePanes to each sheet, configures JsonSaveOptions (AlwaysExportAsJsonObject, ExportNestedStructure, SkipEmptyRows) and saves the workbook as "WorkbookFrozenState.json". The JSON file captures the frozen‑pane layout for auditing or downstream processing.
// Keywords: Aspose.Cells export freeze panes JSON | C# serialize worksheet freeze state | JsonSaveOptions frozen pane example | audit Excel freeze panes programmatically | Aspose.Cells .NET JSON snapshot | freeze panes to JSON file | worksheet layout serialization | Aspose.Cells GitHub sample | Excel freeze pane metadata
// Common Searches: how to export freeze pane settings to JSON using Aspose.Cells | Aspose.Cells C# save worksheet frozen rows and columns as JSON | serialize Excel freeze panes for audit log | Aspose.Cells JsonSaveOptions include frozen pane information | C# example for exporting worksheet layout to JSON
// Developer Intent: Generate a JSON file that records the freeze‑pane configuration of each worksheet for auditing or integration purposes.
// Use Cases: Create an audit trail of freeze‑pane rows and columns across all sheets in a workbook. | Compare worksheet layout before and after automated transformations by diffing JSON snapshots. | Send worksheet structure, including frozen panes, to a web API as a JSON payload.
// AI Prompts: Write C# code that reads the "WorkbookFrozenState.json" produced by Aspose.Cells and extracts the frozen‑pane row and column indices for each worksheet. | Show how to modify the frozen‑pane settings of a worksheet, re‑export the JSON, and verify the changes using Aspose.Cells. | Explain how to configure JsonSaveOptions to output only worksheet metadata (e.g., frozen panes) while excluding cell values.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Json;

// C# example that creates a workbook with two worksheets, applies FreezePanes to each sheet, configures JsonSaveOptions (AlwaysExportAsJsonObject, ExportNestedStructure, SkipEmptyRows) and saves the workbook as "WorkbookFrozenState.json". The JSON file captures the frozen‑pane layout for auditing or downstream processing.
class SerializeFrozenState
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Configure first worksheet with frozen panes
            // -------------------------------------------------
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "FirstSheet";

            // Freeze rows above row 2 (index 1) and columns left of column C (index 2)
            // The last two parameters (totalRows, totalColumns) are set to 0 to freeze all rows/columns before the specified cell
            sheet1.FreezePanes(1, 2, 0, 0);

            // Add some sample data (optional, helps verify JSON output)
            sheet1.Cells["A1"].PutValue("Header1");
            sheet1.Cells["B1"].PutValue("Header2");
            sheet1.Cells["A2"].PutValue("Value1");
            sheet1.Cells["B2"].PutValue(100);

            // -------------------------------------------------
            // Add a second worksheet and also freeze panes
            // -------------------------------------------------
            Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");

            // Freeze rows above row 4 (index 3) and column A (index 1)
            sheet2.FreezePanes(3, 1, 0, 0);

            sheet2.Cells["A1"].PutValue("Item");
            sheet2.Cells["B1"].PutValue("Quantity");
            sheet2.Cells["A2"].PutValue("Apple");
            sheet2.Cells["B2"].PutValue(50);

            // -------------------------------------------------
            // Set JSON save options to export the workbook as a JSON object
            // -------------------------------------------------
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                AlwaysExportAsJsonObject = true, // Export as an object even if only one sheet
                ExportNestedStructure = true,    // Keep hierarchical structure
                SkipEmptyRows = true             // Omit empty rows for cleaner output
            };

            // -------------------------------------------------
            // Save the workbook's frozen state information to a JSON file
            // -------------------------------------------------
            string jsonFilePath = "WorkbookFrozenState.json";

            // Ensure the directory exists (handle case where Path.GetDirectoryName returns null)
            string directory = Path.GetDirectoryName(jsonFilePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            workbook.Save(jsonFilePath, jsonOptions);
            Console.WriteLine($"Frozen state information saved to: {jsonFilePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
