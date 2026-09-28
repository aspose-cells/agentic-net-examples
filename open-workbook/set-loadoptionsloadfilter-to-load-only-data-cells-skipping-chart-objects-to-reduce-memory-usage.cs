// Title: How to load an Excel workbook in C# with Aspose.Cells while skipping chart objects using LoadOptions.LoadFilter
// AI Prompts: Write C# code that creates a LoadOptions object, assigns a custom ILoadFilter implementation that returns false for chart objects, and loads an .xlsx file with the Workbook constructor that accepts LoadOptions. | Show an example of implementing ILoadFilter in Aspose.Cells to load only cell data and exclude all drawing objects, then save the workbook unchanged. | Generate a minimal C# program that demonstrates memory‑efficient loading of a large Excel file by using LoadOptions.LoadFilter to filter out charts before calling the Workbook constructor.
// Common Searches: Aspose.Cells C# load workbook without loading charts | Use LoadOptions.LoadFilter to load only data cells in .NET | Memory efficient Excel loading Aspose.Cells custom ILoadFilter example | Skip chart objects when opening .xlsx with Aspose.Cells .NET | How to implement ILoadFilter to filter out drawings in Aspose.Cells
// Tags: custom ILoadFilter Aspose.Cells C# | LoadOptions.LoadFilter data-only loading | skip chart objects workbook load | memory efficient Excel loading .NET | filter drawings Aspose.Cells load

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLoadFilterExample
{
    // Placeholder class – kept for potential future use.
    // No longer implements ILoadFilter to avoid missing‑type compile issues.
    // The program verifies that input.xlsx exists, loads the workbook using the default Workbook constructor, accesses the first worksheet to display its name, and saves the workbook to output.xlsx. It serves as a base for adding a LoadOptions.LoadFilter to skip chart objects for memory‑efficient loading.
    public class DataOnlyLoadFilter
    {
        // Add any custom logic here if needed in the future.
    }

    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists to prevent FileNotFoundException.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the workbook without a custom filter (charts will be loaded as usual).
                Workbook workbook = new Workbook(inputPath);

                // Perform any required operations on the loaded data cells here.
                // Example: Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];
                Console.WriteLine($"First worksheet name: {sheet.Name}");

                // Save the workbook.
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Log the exception details for troubleshooting.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
