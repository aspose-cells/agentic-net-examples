// Title: Export worksheet names and Tab IDs to JSON with Aspose.Cells for .NET
// Description: Loads an existing workbook, generates a helper sheet that records each worksheet’s name and its index (used as a Tab ID), configures JsonSaveOptions, and saves the result as a formatted JSON file. Perfect for external systems that need a lightweight manifest of sheet identifiers.
// Keywords: Aspose.Cells | C# | .NET | export worksheet names to JSON | worksheet Tab ID | sheet index JSON | JsonSaveOptions example | Excel workbook metadata | configuration file generation | external integration
// Common Searches: Aspose.Cells export sheet names to JSON | C# get worksheet index as TabId Aspose.Cells | How to save workbook metadata as JSON using Aspose.Cells | Create JSON manifest of Excel sheets .NET | JsonSaveOptions usage Aspose.Cells
// Developer Intent: Generate a JSON file that maps every worksheet name to its internal Tab ID (index) for consumption by external applications.
// Use Cases: Provide a JSON lookup for a reporting engine that references worksheets by index. | Synchronize UI navigation with Excel workbooks by exposing sheet names and IDs to a web service. | Automate data import scripts that require a manifest of sheet identifiers without opening the workbook.
// AI Prompts: Write C# code using Aspose.Cells to export all worksheet names and their indexes to a JSON file with a header row. | Show how to configure JsonSaveOptions to produce indented JSON and omit empty cells. | Explain how to obtain a worksheet’s Tab ID in Aspose.Cells when a direct property is unavailable.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Loads an existing workbook, generates a helper sheet that records each worksheet’s name and its index (used as a Tab ID), configures JsonSaveOptions, and saves the result as a formatted JSON file. Perfect for external systems that need a lightweight manifest of sheet identifiers.
class ExportTabIdToJson
{
    static void Main()
    {
        // Load the source workbook (replace with actual path)
        Workbook sourceWorkbook = new Workbook("input.xlsx");

        // Create a new workbook to store TabId information
        Workbook tabIdWorkbook = new Workbook();
        // Remove the default sheet and add a dedicated sheet
        tabIdWorkbook.Worksheets.Clear();
        Worksheet tabIdSheet = tabIdWorkbook.Worksheets.Add("TabIds");

        // Write header row
        tabIdSheet.Cells["A1"].PutValue("SheetName");
        tabIdSheet.Cells["B1"].PutValue("TabId");

        // Populate TabId data for each worksheet in the source workbook
        for (int i = 0; i < sourceWorkbook.Worksheets.Count; i++)
        {
            Worksheet ws = sourceWorkbook.Worksheets[i];
            int rowIndex = i + 1; // zero‑based index; row 2 onward

            // Sheet name
            tabIdSheet.Cells[rowIndex, 0].PutValue(ws.Name);

            // TabId – using worksheet index as a fallback (Aspose.Cells does not expose a TabId property directly)
            int tabId = ws.Index;
            tabIdSheet.Cells[rowIndex, 1].PutValue(tabId);
        }

        // Configure JSON save options
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            HasHeaderRow = true,          // treat first row as header
            ExportNestedStructure = false,
            ExportEmptyCells = false,
            ExportAsString = false,
            Indent = "  "
        };

        // Save the TabId workbook as a JSON configuration file
        string outputPath = "WorksheetTabIds.json";
        tabIdWorkbook.Save(outputPath, jsonOptions);
    }
}
