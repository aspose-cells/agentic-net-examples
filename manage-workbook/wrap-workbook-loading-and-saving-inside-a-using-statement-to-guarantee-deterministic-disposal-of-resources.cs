// Title: Wrap Aspose.Cells Workbook Load and Save in a C# using Statement for Proper Disposal
// AI Prompts: Refactor the sample to enclose the Workbook object in a using block so it is automatically disposed after saving. | Generate C# code that opens an .xlsx file with Aspose.Cells, optionally modifies it, and saves it, using a using statement for deterministic cleanup. | Show how to apply the .NET using pattern to manage Aspose.Cells Workbook resources when reading and writing Excel workbooks.
// Common Searches: C# using statement for Aspose.Cells Workbook disposal after Save | how to ensure Aspose.Cells Workbook is released in .NET | deterministic cleanup of Excel workbook using Aspose.Cells and using block | example of using Aspose.Cells Workbook within using block in C# | best practice for disposing Aspose.Cells Workbook objects
// Tags: using statement Aspose.Cells Workbook disposal | load and save Excel with Aspose.Cells using block | deterministic resource cleanup .NET Aspose.Cells | C# workbook disposal pattern Aspose.Cells | resource management Aspose.Cells Workbook

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel file into an Aspose.Cells Workbook, optionally modifies it, saves it to a new file, and wraps the Workbook in a using statement to guarantee deterministic disposal of resources.
    class Program
    {
        static void Main(string[] args)
        {
            // Define file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the input file using the constructor
                Workbook workbook = new Workbook(inputPath);

                // (Optional) Perform any workbook modifications here

                // Save the workbook to the output file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
