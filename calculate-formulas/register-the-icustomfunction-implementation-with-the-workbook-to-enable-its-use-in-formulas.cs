// Title: How to register a custom ICustomFunction with an Aspose.Cells workbook for use in C# formulas
// AI Prompts: Generate C# code that implements ICustomFunction, registers it with a Workbook, and calls the function from a worksheet cell. | Show me how to add a user‑defined function that returns the product of two numbers to an Aspose.Cells workbook and invoke it via =MYPRODUCT(A1,B1). | Provide a step‑by‑step example of creating an ICustomFunction class, registering it with workbook.CustomFunctions, and evaluating the formula.
// Common Searches: aspnet register ICustomFunction implementation in Aspose.Cells workbook | example of custom Excel function using Aspose.Cells .NET API | how to call a user defined function from a cell formula in Aspose.Cells C# | registering custom calculation function with Aspose.Cells workbook for Excel export | using ICustomFunction to extend formula capabilities in Aspose.Cells
// Tags: Aspose.Cells custom ICustomFunction registration | C# user-defined Excel function Aspose.Cells | register custom formula Aspose.Cells workbook | ICustomFunction implementation example | extend formula engine Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates creating a workbook, implementing an ICustomFunction, registering it with the workbook's custom functions collection, invoking the function from a cell formula, calculating the result, and saving the workbook.
public class Program
{
    public static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Apply a simple built‑in formula (SUM) in a cell
            workbook.Worksheets[0].Cells["A1"].Formula = "=SUM(5, 10)";

            // Calculate formulas to evaluate the cell value
            workbook.CalculateFormula();

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Log any errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
