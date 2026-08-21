// Title: Load XLSB, Recalculate All Formulas, and Export to JSON with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to open an XLSB workbook using Aspose.Cells LoadOptions, force a full formula recalculation with CalculateFormula, and save the result as JSON via SaveFormat.Json into a MemoryStream for further processing.
// Keywords: Aspose.Cells XLSB load C# | recalculate formulas Aspose.Cells | export workbook to JSON | SaveFormat.Json Aspose.Cells | convert binary Excel to JSON | .NET Excel to JSON conversion | memory stream JSON Aspose
// Common Searches: Aspose.Cells read XLSB and output JSON | C# calculate formulas before exporting Excel to JSON | How to convert XLSB to JSON with Aspose.Cells | Save Excel workbook as JSON in .NET | Recalculate all formulas in XLSB using Aspose
// Developer Intent: Produce a JSON representation of an XLSB workbook after evaluating every formula.
// Use Cases: Transform an XLSB financial model into JSON for a web API after all calculations are up‑to‑date. | Batch‑process multiple XLSB reports, recalculate values, and feed the JSON output into a data‑analytics pipeline. | Expose spreadsheet data to a JavaScript front‑end by converting the fully evaluated workbook to JSON.
// AI Prompts: Show how to limit the JSON export to a single worksheet. | Provide code to stream the JSON directly to a file instead of printing to the console. | Suggest performance‑optimizing techniques for converting large XLSB files to JSON with Aspose.Cells.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// Demonstrates how to open an XLSB workbook using Aspose.Cells LoadOptions, force a full formula recalculation with CalculateFormula, and save the result as JSON via SaveFormat.Json into a MemoryStream for further processing.
class Program
{
    static void Main()
    {
        // Path to the source XLSB file
        string xlsbPath = "input.xlsb";

        // Load the XLSB workbook with appropriate load options
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsb);
        Workbook workbook = new Workbook(xlsbPath, loadOptions);

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Save the workbook to a memory stream in JSON format
        using (MemoryStream jsonStream = new MemoryStream())
        {
            workbook.Save(jsonStream, SaveFormat.Json);
            jsonStream.Seek(0, SeekOrigin.Begin);
            string jsonResult = Encoding.UTF8.GetString(jsonStream.ToArray());

            // Output the JSON representation
            Console.WriteLine(jsonResult);
        }

        // Clean up
        workbook.Dispose();
    }
}
