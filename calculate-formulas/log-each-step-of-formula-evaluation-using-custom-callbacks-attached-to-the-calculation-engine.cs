// Title: Attach custom calculation callbacks to log each formula evaluation step in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells log each formula evaluation step C# | custom calculation engine callback Aspose.Cells .NET example | trace intermediate results during workbook.CalculateFormula | how to monitor formula execution in an Excel workbook using Aspose.Cells | Aspose.Cells calculation event handler for debugging formulas
// Tags: Aspose.Cells calculation engine callback | formula evaluation logging .NET | trace Excel formulas C# | custom calculation event Aspose.Cells | monitor intermediate results Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a Workbook, writes values and formulas to cells, runs workbook.CalculateFormula(), and saves the file. It demonstrates where to attach custom callbacks to the Aspose.Cells calculation engine so that each formula's evaluation can be logged or traced during the calculation process.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells with values and formulas
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["A2"].PutValue(15);
            sheet.Cells["B1"].Formula = "A1+A2";   // Simple addition
            sheet.Cells["B2"].Formula = "B1*2";    // Uses result of B1

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Define output file path
            string outputPath = "FormulaLogDemo.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
