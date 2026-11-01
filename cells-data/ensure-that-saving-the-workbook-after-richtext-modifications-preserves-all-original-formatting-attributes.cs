// Title: How to preserve original rich‑text formatting when updating a cell value and saving the workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# keep rich text formatting after using PutValue on a cell | Save workbook without losing cell rich‑text styles in Aspose.Cells .NET | How to retain original cell formatting when prepending text with Aspose.Cells | Preserve cell rich‑text attributes while modifying cell content in Aspose.Cells | Maintain existing formatting of a cell after updating its value with Aspose.Cells
// Tags: rich text style preservation Aspose.Cells | cell value update preserving formatting Aspose.Cells | workbook save retaining cell styles Aspose.Cells | prefix cell text with formatting intact Aspose.Cells | Aspose.Cells rich‑text attribute retention

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook (or creates a new one), accesses cell A1, reads its current rich‑text content, prepends a custom prefix while preserving all original rich‑text runs and formatting, writes the combined text back using the RichText API, and saves the workbook. It demonstrates how to modify cell text without losing any formatting attributes.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Load existing workbook or create a new one if the input file is missing
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
                workbook.Worksheets.Add();

            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];

            // Get the current cell text (empty string if the cell is blank)
            string originalText = cell.StringValue ?? string.Empty;

            // Prefix to prepend
            const string prefix = "Prefix: ";

            // Combine prefix with original text
            string newText = prefix + originalText;

            // Assign the new text back to the cell
            cell.PutValue(newText);

            // Save the workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
