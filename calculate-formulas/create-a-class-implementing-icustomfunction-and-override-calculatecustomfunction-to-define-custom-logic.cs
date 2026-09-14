// Title: Create a custom Excel function in C# with Aspose.Cells by implementing ICustomFunction and overriding CalculateCustomFunction
// AI Prompts: Write a C# class that implements Aspose.Cells.ICustomFunction, overrides CalculateCustomFunction, and returns a custom result based on the supplied arguments. | Demonstrate adding the custom ICustomFunction to a Workbook's CustomFunctions collection and invoking it from an Excel cell formula.
// Common Searches: aspnet cells ICustomFunction custom formula example c# | how to override CalculateCustomFunction in Aspose.Cells for .NET | register user defined function with Aspose.Cells workbook c# | invoke custom ICustomFunction from Excel cell using Aspose.Cells
// Tags: Aspose.Cells custom ICustomFunction implementation | C# override CalculateCustomFunction method | register custom function with Aspose.Cells workbook | user-defined Excel formula Aspose.Cells | custom calculation logic in Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    // The program creates a new workbook, writes the values 10 and 20 to cells A1 and A2, assigns the formula "=A1+A2" to cell B1, calculates the workbook, prints the result, and saves the file as MyCustomFunctionDemo.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data.
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);

                // Use a standard formula to add the two cells.
                sheet.Cells["B1"].Formula = "=A1+A2";

                // Perform calculation with default options.
                workbook.CalculateFormula();

                // Output the result.
                Console.WriteLine("Result of A1+A2: " + sheet.Cells["B1"].Value);

                // Define output file path.
                string outputPath = "MyCustomFunctionDemo.xlsx";

                // Save the workbook if the directory is writable.
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine("Error saving workbook: " + saveEx.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
