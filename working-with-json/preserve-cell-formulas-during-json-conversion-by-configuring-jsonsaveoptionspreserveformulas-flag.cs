// Title: Export Workbook to JSON with Formulas Using Aspose.Cells JsonSaveOptions in C#
// Description: Demonstrates how to save an Aspose.Cells workbook as a JSON file while keeping cell formulas intact by configuring the JsonSaveOptions.PreserveFormulas flag (enabled by default). The example creates a workbook, adds numeric values and a SUM formula, sets the JSON save options, and writes the file.
// Keywords: Aspose.Cells | C# | JsonSaveOptions | PreserveFormulas | export Excel to JSON | save workbook as JSON | keep formulas in JSON | Excel formulas JSON conversion | Aspose.Cells JSON example | C# Excel to JSON
// Common Searches: Aspose.Cells preserve formulas when saving to JSON | JsonSaveOptions PreserveFormulas C# example | Export Excel workbook to JSON with formulas | How to keep formulas in JSON using Aspose.Cells | C# save workbook as JSON without losing formulas
// Developer Intent: Create a JSON representation of an Excel workbook that retains all cell formulas.
// Use Cases: Send a financial model to a web API in JSON while preserving calculation logic. | Synchronize spreadsheet data with a JavaScript front‑end that evaluates formulas client‑side. | Archive Excel worksheets in JSON format for version control without stripping formulas.
// AI Prompts: Write C# code that uses Aspose.Cells JsonSaveOptions with PreserveFormulas enabled to export a workbook to JSON. | Explain how JsonSaveOptions.PreserveFormulas influences the JSON output and compare results with and without the flag. | Generate a C# snippet that reads a JSON file saved with formulas and loads it back into a new Aspose.Cells workbook, keeping the formulas intact.

using System;
using Aspose.Cells;

// Demonstrates how to save an Aspose.Cells workbook as a JSON file while keeping cell formulas intact by configuring the JsonSaveOptions.PreserveFormulas flag (enabled by default). The example creates a workbook, adds numeric values and a SUM formula, sets the JSON save options, and writes the file.
class PreserveFormulasJsonDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate cells with values and a formula
            worksheet.Cells["A1"].PutValue(10);
            worksheet.Cells["A2"].PutValue(20);
            worksheet.Cells["A3"].Formula = "=SUM(A1:A2)";

            // Configure JSON save options (formulas are exported by default)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Save the workbook as a JSON file using the configured options
            workbook.Save("preserve_formulas.json", jsonOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
