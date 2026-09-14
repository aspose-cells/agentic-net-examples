// Title: How to read the evaluated value of cell D5 after recalculating formulas with Aspose.Cells in C#
// AI Prompts: Load an Excel workbook with Aspose.Cells, invoke Workbook.CalculateFormula, and return the value of cell D5 as a C# object. | Adapt the code to accept any cell address and output its calculated value after Worksheet.CalculateFormula execution. | Format the retrieved D5 value to two decimal places and display it on the console using C#.
// Common Searches: Aspose.Cells C# get value of a cell after calling CalculateFormula | C# read evaluated result of Excel cell D5 using Aspose.Cells | How to recalculate formulas and fetch a specific cell value with Aspose.Cells for .NET | Example of Workbook.CalculateFormula and accessing cell value in C#
// Tags: aspocells calculateformula read cell value | c# retrieve evaluated excel cell aspocells | worksheet recalculate formulas get cell result | excel cell d5 value after recalculation aspocells | load workbook calculate all formulas aspocells

using Aspose.Cells;
using System;
using System.IO;

// Loads 'input.xlsx' with Aspose.Cells, recalculates all formulas using Workbook.CalculateFormula, reads the evaluated value of cell D5, and prints it to the console.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Get the first worksheet (or specify by name/index)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve cell D5 (row 4, column 3)
            Cell cellD5 = worksheet.Cells["D5"];

            // Read the calculated value
            object calculatedValue = cellD5.Value;

            // Display the result
            Console.WriteLine("Calculated value of D5: " + calculatedValue);
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
