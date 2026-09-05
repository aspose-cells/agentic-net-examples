// Title: Protect an Excel worksheet with Aspose.Cells for .NET while allowing sorting and disabling filtering
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, applies password protection to the first worksheet, enables sorting, disables filtering, and saves the workbook. | Demonstrate how to configure worksheet protection settings in Aspose.Cells for .NET to permit sorting operations but block filter usage.
// Common Searches: Aspose.Cells .NET protect worksheet allow sorting disable filtering example | C# set worksheet protection to enable sort only using Aspose.Cells | How to apply password protection to an Excel sheet while permitting sorting in Aspose.Cells
// Tags: Aspose.Cells worksheet protection allow sorting | disable filtering on protected worksheet C# | password protect Excel sheet Aspose.Cells .NET | configure worksheet protection options Aspose.Cells | Excel sorting permission Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads an existing workbook, accesses the first worksheet, configures its protection to allow sorting but prevent filtering, applies a password using ProtectionType.All, and saves the modified file, handling missing input files and runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (index 0)
                Worksheet sheet = workbook.Worksheets[0];

                // Allow sorting but disallow filtering
                sheet.Protection.AllowSorting = true;
                sheet.Protection.AllowFiltering = false;

                // Protect the worksheet with a password (oldPassword not required, pass null)
                sheet.Protect(ProtectionType.All, "YourPassword", null);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display the message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
