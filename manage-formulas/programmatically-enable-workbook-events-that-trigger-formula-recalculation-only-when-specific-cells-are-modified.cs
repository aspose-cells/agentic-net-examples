// Title: Recalculating workbook formulas only after specific cell updates using Aspose.Cells for .NET
// Common Searches: Aspose.Cells recalculate formulas after changing cell A3 in C# | C# trigger workbook.CalculateFormula only when certain cells are edited with Aspose.Cells | How to limit formula recalculation to selected cells in Aspose.Cells .NET | Programmatically invoke formula calculation after specific cell modification using Aspose.Cells | Selective formula update in Excel file using Aspose.Cells C# API
// Tags: selective formula recalculation Aspose.Cells .NET | manual workbook.CalculateFormula invocation C# | cell-specific formula update Aspose.Cells | Excel workbook event simulation Aspose.Cells | conditional formula evaluation C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, adds a SUM formula in cell A6 that references A1:A5, modifies cell A3, manually calls workbook.CalculateFormula() to recalculate dependent formulas, and saves the workbook as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Example formula that depends on cells A1:A5
            sheet.Cells["A6"].Formula = "=SUM(A1:A5)";

            // Change a cell that should cause recalculation (A3)
            sheet.Cells["A3"].PutValue(15);

            // Manually recalculate formulas after the change
            workbook.CalculateFormula();

            // Define output file path
            string outputPath = "Result.xlsx";

            // Ensure the directory exists (if any)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
