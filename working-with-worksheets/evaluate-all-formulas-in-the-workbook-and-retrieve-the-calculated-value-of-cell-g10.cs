using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the workbook from a file (replace with your actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Evaluate all formulas in the entire workbook
        workbook.CalculateFormula();

        // Access cell G10 in the first worksheet and get its calculated value
        Cell g10Cell = workbook.Worksheets[0].Cells["G10"];
        object g10Value = g10Cell.Value;

        // Display the result
        Console.WriteLine("Calculated value of G10: " + g10Value);

        // Save the workbook after calculation (optional)
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}