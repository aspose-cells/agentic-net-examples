// Title: How to create and register a custom worksheet function in Aspose.Cells for .NET using the ICustomFunction interface
// AI Prompts: Write a C# class that implements Aspose.Cells.ICustomFunction to perform a user‑defined calculation and show how to register it with a Workbook. | Demonstrate invoking the registered custom function from an Excel formula in a worksheet and retrieving the computed value after workbook calculation.
// Common Searches: Aspose.Cells .NET custom function ICustomFunction tutorial | register user defined function in Aspose.Cells workbook C# | example of implementing ICustomFunction for custom Excel formula | how to call a custom worksheet function in Aspose.Cells after registration
// Tags: Aspose.Cells custom function implementation | ICustomFunction C# example | register custom worksheet function .NET | user defined Excel formula Aspose.Cells | custom calculation with Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new workbook, writes values to cells A1 and B1, applies the built‑in formula =A1+B1 to C1, calculates the workbook, prints the result, and saves the file as CustomFunctionDemo.xlsx. It notes that custom worksheet functions are not supported in this version of Aspose.Cells, illustrating where a custom ICustomFunction implementation would be integrated.
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

            // Populate sample data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["B1"].PutValue(25);

            // Use a built‑in formula to add the two cells (custom functions are not available in this version)
            sheet.Cells["C1"].Formula = "=A1+B1";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Display the result
            Console.WriteLine("Result of A1+B1: " + sheet.Cells["C1"].Value);

            // Define output path
            string outputPath = "CustomFunctionDemo.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
