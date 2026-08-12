// Title: Export Active Worksheet to JSON with Aspose.Cells for .NET (default JsonSaveOptions)
// Description: Loads an Excel workbook, identifies the active worksheet, configures JsonSaveOptions with default settings, restricts the export to that sheet via SheetIndexes, and saves the result as a JSON file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | export active worksheet to JSON | JsonSaveOptions | SheetIndexes | Excel to JSON conversion | workbook.Save JSON | default save format | code example
// Common Searches: Aspose.Cells export active sheet JSON C# | How to save only the active worksheet as JSON using Aspose.Cells | JsonSaveOptions default settings example | Convert Excel active tab to JSON .NET | C# code to export specific worksheet to JSON
// Developer Intent: Generate a JSON file that contains only the currently active worksheet of a loaded Excel workbook using Aspose.Cells with default save options.
// Use Cases: Provide a lightweight JSON payload of the user‑selected sheet for a web API. | Create front‑end data feeds by converting the active Excel tab to JSON. | Automate server‑side reporting that requires only the active worksheet in JSON format.
// AI Prompts: Write C# code using Aspose.Cells to export the active worksheet of a workbook to JSON with default JsonSaveOptions. | Show how to limit Aspose.Cells JSON export to a single sheet by setting the SheetIndexes property. | Explain how to adapt the example to export multiple worksheets to JSON by modifying SheetIndexes.

using System;
using Aspose.Cells;

// Loads an Excel workbook, identifies the active worksheet, configures JsonSaveOptions with default settings, restricts the export to that sheet via SheetIndexes, and saves the result as a JSON file using Aspose.Cells for .NET.
class ExportActiveWorksheetToJson
{
    static void Main()
    {
        // Load the workbook from a file (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Determine the index of the active worksheet
        int activeSheetIndex = workbook.Worksheets.ActiveSheetIndex;

        // Create JSON save options with default settings
        JsonSaveOptions jsonOptions = new JsonSaveOptions();

        // Restrict export to only the active worksheet
        jsonOptions.SheetIndexes = new int[] { activeSheetIndex };

        // Save the active worksheet as a JSON file
        workbook.Save("output.json", jsonOptions);
    }
}
