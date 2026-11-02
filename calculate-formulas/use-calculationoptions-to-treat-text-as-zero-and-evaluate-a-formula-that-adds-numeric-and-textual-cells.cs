// Title: Using Aspose.Cells CalculationOptions.TreatTextAsZero in C# to add numeric and text cells
// Common Searches: Aspose.Cells C# set CalculationOptions.TreatTextAsZero example | Calculate Excel formula with text values ignored using Aspose.Cells | C# add a numeric cell and a text numeric string with Aspose.Cells | How to configure Aspose.Cells to treat text as zero during formula evaluation | Aspose.Cells workbook.CalculateFormula with TreatTextAsZero C#
// Tags: Aspose.Cells calculation options API C# | Ignore text values in Excel formula Aspose.Cells | Add numeric and text cells with Aspose.Cells | Workbook.CalculateFormula configuration Aspose.Cells | Excel calculation treating text as zero Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, places the number 10 in cell A1 and the text "20" in cell A2, assigns the formula "=A1+A2" to A3, uses the default CalculationOptions that treat text as zero, recalculates the worksheet, reads the numeric result from A3, and saves the workbook to a file.
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

            // Numeric value in A1
            sheet.Cells["A1"].PutValue(10);

            // Textual numeric value in A2
            sheet.Cells["A2"].PutValue("20");

            // Formula in A3 that adds A1 and A2
            sheet.Cells["A3"].Formula = "=A1+A2";

            // Recalculate formulas (default behavior treats text as zero)
            workbook.CalculateFormula();

            // Retrieve the result from A3
            double result = sheet.Cells["A3"].DoubleValue;
            Console.WriteLine("Result of A3: " + result);

            // Save the workbook (optional)
            string outputPath = "Result.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
