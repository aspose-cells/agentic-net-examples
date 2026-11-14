// Title: Refresh linked shapes after modifying source cells using Aspose.Cells for .NET in C#
// Common Searches: Aspose.Cells C# refresh chart linked to cell after changing cell value | Update picture linked to worksheet cell and recalculate workbook with Aspose.Cells | How to recalculate formulas to refresh linked shapes in an Excel file using Aspose.Cells .NET | C# code to refresh linked shapes after modifying source cells in Aspose.Cells | Aspose.Cells recalculate workbook to update linked objects after cell update
// Tags: refresh linked shapes Aspose.Cells .NET | recalculate workbook after cell update C# | update chart data source cells Aspose.Cells | linked picture refresh Aspose.Cells | modify cell values and refresh shapes C#

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, changes the values of cells A1 and B2, calls CalculateFormula to recalculate formulas and refresh any shapes (charts, pictures, etc.) linked to those cells, and then saves the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (or specify the needed one)
            Worksheet worksheet = workbook.Worksheets[0];

            // Update the source cells that linked shapes depend on
            worksheet.Cells["A1"].PutValue(123);   // example value
            worksheet.Cells["B2"].PutValue(456);   // example value

            // Recalculate formulas so that linked shapes (charts, pictures, etc.) are refreshed
            workbook.CalculateFormula();

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
