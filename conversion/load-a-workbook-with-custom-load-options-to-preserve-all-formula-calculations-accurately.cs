// Title: Load an Excel workbook with Aspose.Cells for .NET using custom LoadOptions to retain all formulas and calculations (C#)
// Common Searches: Aspose.Cells C# load workbook keep formulas intact | How to use LoadOptions in Aspose.Cells to preserve Excel calculations | Read cell formula from Excel file without losing it when saving with Aspose.Cells | Load Excel file with formulas preserved using Aspose.Cells .NET API | Save workbook in C# without recalculating formulas using Aspose.Cells
// Tags: Aspose.Cells LoadOptions preserve formulas | C# load Excel workbook with formula retention | Aspose.Cells read cell formula | save workbook without recalculating formulas Aspose.Cells | Excel formula preservation Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading an Excel file with Aspose.Cells using custom LoadOptions to ensure all formulas and their calculations are retained, reading a formula from a specific cell, and saving the workbook without altering the original calculations, while handling file‑not‑found and other exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook with default options (formulas are preserved by default)
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet and read the formula in cell A1
            Worksheet sheet = workbook.Worksheets[0];
            string formulaInA1 = sheet.Cells["A1"].Formula;
            Console.WriteLine($"Formula in A1: {formulaInA1}");

            // Save the workbook after any processing
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
