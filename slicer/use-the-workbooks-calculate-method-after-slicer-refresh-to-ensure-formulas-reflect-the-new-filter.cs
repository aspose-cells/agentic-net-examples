// Title: Refresh a slicer and recalculate all formulas in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Load an .xlsx file, call Worksheet.Slicers[0].Refresh(), invoke Workbook.CalculateFormula(), and save the updated workbook with Aspose.Cells in C#. | Programmatically refresh an Excel slicer and trigger full workbook formula recalculation using the Aspose.Cells API. | Use Aspose.Cells to apply a slicer refresh, recalculate dependent formulas, and export the modified workbook.
// Common Searches: Aspose.Cells how to refresh slicer and recalculate formulas in C# | C# programmatically refresh Excel slicer using Aspose.Cells | Recalculate workbook after slicer change with Aspose.Cells .NET | Example of Worksheet.Slicers[0].Refresh() and Workbook.CalculateFormula()
// Tags: Aspose.Cells slicer refresh C# | Workbook.CalculateFormula usage Aspose.Cells | Excel slicer programmatic update .NET | Recalculate dependent formulas after slicer Aspose.Cells | Load and save workbook with slicer refresh Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel file, refreshes the first slicer on the first worksheet, recalculates all formulas with CalculateFormula, and saves the updated workbook to a new file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume the slicer is on the first worksheet and is the first slicer
            Worksheet worksheet = workbook.Worksheets[0];
            if (worksheet.Slicers.Count > 0)
            {
                // Refresh the slicer to apply the current filter
                worksheet.Slicers[0].Refresh();
            }

            // Recalculate all formulas after the slicer refresh
            workbook.CalculateFormula();

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
