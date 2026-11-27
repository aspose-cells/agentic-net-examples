// Title: Use Aspose.Cells for .NET to evaluate all formulas and obtain the calculated value of cell G10 in an Excel worksheet
// Common Searches: Aspose.Cells C# evaluate all formulas in an Excel workbook and get the result of cell G10 | How to calculate formulas and read the string value of a specific cell using Aspose.Cells .NET | C# code to load an .xlsx file, run Workbook.CalculateFormula, and retrieve G10 value | Retrieve evaluated cell value after calling CalculateFormula with Aspose.Cells in C# | Get the computed result of G10 after formula calculation in Aspose.Cells for .NET
// Tags: Workbook.CalculateFormula Aspose.Cells | read cell stringvalue Aspose.Cells C# | evaluate all Excel formulas .NET | load .xlsx and compute formulas Aspose.Cells | access cell G10 result C#

using System;
using Aspose.Cells;

namespace EvaluateFormulasExample
{
    // Demonstrates loading an .xlsx workbook with Aspose.Cells for .NET, invoking Workbook.CalculateFormula to evaluate every formula, accessing the first worksheet, and reading the StringValue of cell G10 to obtain the computed result.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the input Excel file. Adjust the path as needed.
            const string inputFilePath = "input.xlsx";

            // Verify that the file exists before attempting to load it.
            if (!System.IO.File.Exists(inputFilePath))
            {
                Console.WriteLine($"Error: The file \"{inputFilePath}\" was not found.");
                return;
            }

            try
            {
                // Load the workbook from the specified file.
                Workbook workbook = new Workbook(inputFilePath);

                // Calculate all formulas in the entire workbook.
                // This method evaluates every formula cell across all worksheets.
                workbook.CalculateFormula();

                // Access the first worksheet (index 0). Change the index if the target cell is on another sheet.
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the calculated value of cell G10.
                // The Value property returns the underlying object (double, string, DateTime, etc.).
                // For display purposes, we use the StringValue which formats the value as a string.
                Cell targetCell = worksheet.Cells["G10"];
                string calculatedValue = targetCell.StringValue;

                Console.WriteLine($"Calculated value of G10: {calculatedValue}");
            }
            catch (Exception ex)
            {
                // Output any exception details to aid troubleshooting.
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
