// Title: How to retrieve and log the data validation type and formulas of cell J5 with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel workbook, accesses cell J5, obtains its Validation object, and prints Validation.Type, Validation.Formula1, and Validation.Formula2 to the console using Aspose.Cells. | Show how to verify whether a specific cell has data validation and output its type and formula values in a .NET console application with Aspose.Cells.
// Common Searches: Aspose.Cells C# read data validation type from a specific cell | Get validation formulas for a target cell using Aspose.Cells .NET | How to determine if a cell contains data validation with Aspose.Cells | Log Excel validation details in a .NET console application
// Tags: aspocells read cell validation .net | aspocells retrieve validation formulas c# | excel data validation extraction using aspocells | c# console log validation type aspocells | target cell validation aspocells example

using System;
using Aspose.Cells;

// // Loads an Excel workbook, accesses cell J5 on the first worksheet, obtains its Validation object (if present), and writes the validation type together with Formula1 and Formula2 to the console.
class Program
{
    static void Main()
    {
        // Load an existing workbook that contains validation on cell J5.
        // Replace "input.xlsx" with the actual file path.
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or change index as needed).
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the cell J5.
        Cell cell = worksheet.Cells["J5"];

        // Retrieve the validation applied to this cell.
        Validation validation = cell.GetValidation();

        if (validation != null)
        {
            // Log the validation type.
            Console.WriteLine("Validation Type: " + validation.Type);

            // Log the first and second formula values (if any).
            Console.WriteLine("Formula1: " + validation.Formula1);
            Console.WriteLine("Formula2: " + validation.Formula2);
        }
        else
        {
            Console.WriteLine("No validation applied to cell J5.");
        }
    }
}
