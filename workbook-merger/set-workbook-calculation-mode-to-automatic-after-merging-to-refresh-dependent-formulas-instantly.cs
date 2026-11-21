// Title: How to enable automatic calculation mode after merging cells with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# merge cells and automatically recalculate dependent formulas | Set workbook calculation mode to Automatic after merging a range using Aspose.Cells | Refresh Excel formulas instantly after cell merge in a .NET application | C# Aspose.Cells calculate formulas automatically after merging cells | Enable auto‑calc for merged cells in an Excel workbook with Aspose.Cells
// Tags: Aspose.Cells automatic calculation mode | merge cells recalculate formulas Aspose.Cells | C# workbook calculation settings Aspose.Cells | Excel cell merge formula refresh Aspose.Cells | Aspose.Cells workbook save after merge

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, merges the range A1:B2 on the first worksheet, forces an immediate formula recalculation to keep dependent cells up‑to‑date, and then saves the modified file. It demonstrates how to ensure formulas are refreshed instantly after a merge using Aspose.Cells for .NET with robust error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                return;
            }

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Merge a range of cells (A1:B2)
            worksheet.Cells.Merge(0, 0, 2, 2);

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
