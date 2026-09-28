// Title: Recalculate formulas and freeze the first row in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that loads an existing .xlsx file, recalculates all formulas, freezes the top row, and saves the updated workbook. | Demonstrate how to call Worksheet.FreezePanes after Workbook.CalculateFormula to keep refreshed data visible in an Excel file with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# recalculate all formulas and freeze top row in Excel file | How to apply FreezePanes after updating data connections with Aspose.Cells .NET | C# example to load workbook, calculate formulas, and freeze first row using Aspose.Cells | Programmatically freeze header row after refreshing data in Excel via Aspose.Cells
// Tags: Workbook.CalculateFormula Aspose.Cells | Worksheet.FreezePanes top row | load .xlsx Aspose.Cells C# | save workbook with frozen panes Aspose.Cells | fallback data connection refresh Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example verifies the input file, loads it with Aspose.Cells, recalculates all formulas as a fallback for refreshing data connections, freezes the first row using FreezePanes, and saves the modified workbook to the specified output path.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Recalculate all formulas (acts as a fallback when RefreshDataConnections is unavailable)
            workbook.CalculateFormula();

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Freeze the top row to keep refreshed results visible
            // Parameters: total rows to freeze, total columns to freeze, rows, columns
            sheet.FreezePanes(1, 0, 1, 0);

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
