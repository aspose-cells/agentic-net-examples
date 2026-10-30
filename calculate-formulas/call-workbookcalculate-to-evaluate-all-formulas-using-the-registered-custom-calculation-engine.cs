// Title: Calculate all Excel formulas in a workbook with Aspose.Cells CalculateFormula using C#
// Common Searches: how to evaluate every formula in an .xlsx file using Aspose.Cells for .NET | C# code to run workbook.CalculateFormula after loading an Excel workbook | using Aspose.Cells to recalculate formulas with a custom calculation engine in .NET | programmatically recalculate Excel workbook formulas before saving with Aspose.Cells | example of loading, calculating, and saving an Excel file using Aspose.Cells in C#
// Tags: Aspose.Cells CalculateFormula API | C# Excel formula evaluation | custom calculation engine integration Aspose.Cells | load and recalculate workbook Aspose.Cells | save workbook after formula calculation C#

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file (or creates a new workbook), invokes workbook.CalculateFormula() to evaluate all formulas using the registered calculation engine, and then saves the updated workbook, handling any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            // Evaluate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the workbook after calculation
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
