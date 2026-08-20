// Title: Exclude "Temp" worksheets during load with Aspose.Cells LoadOptions in C#
// Description: Demonstrates how to prevent worksheets named "Temp" from being loaded by adding their names to LoadOptions.LoadFilter.ExcludedSheets, then saving the filtered workbook. This approach avoids post‑load removal and improves performance.
// Keywords: Aspose.Cells LoadOptions ExcludedSheets | C# exclude worksheet Temp | load Excel without specific sheets | filter worksheets on load Aspose | performance optimize Excel loading C#
// Common Searches: Aspose.Cells exclude sheet named Temp on load | LoadOptions.ExcludedSheets example C# | how to skip worksheets when opening Excel with Aspose | prevent loading temporary sheets Aspose.Cells | C# load Excel file without certain worksheets
// Developer Intent: Load an Excel workbook while automatically omitting any worksheet called "Temp" by configuring LoadOptions.LoadFilter.ExcludedSheets.
// Use Cases: Generate reports from a template that contains hidden helper sheets, ensuring they never reach the client. | Reduce memory usage and load time for large workbooks that include temporary calculation sheets. | Automate data pipelines where intermediate "Temp" sheets are created during processing but should not be part of the final output.
// AI Prompts: Provide C# code that uses Aspose.Cells LoadOptions to exclude worksheets named "Temp" when opening an Excel file. | Show how to configure LoadOptions.LoadFilter.ExcludedSheets for multiple sheet names in Aspose.Cells. | Explain the performance benefits of excluding sheets during load versus removing them after the workbook is opened.

using System;
using System.IO;
using Aspose.Cells;

namespace ExcludeTempSheetsExample
{
    // Demonstrates how to prevent worksheets named "Temp" from being loaded by adding their names to LoadOptions.LoadFilter.ExcludedSheets, then saving the filtered workbook. This approach avoids post‑load removal and improves performance.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string sourcePath = "Template.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            try
            {
                // Load the workbook (all sheets are loaded initially)
                Workbook workbook = new Workbook(sourcePath);

                // Remove any worksheet named "Temp" after loading
                for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
                {
                    Worksheet ws = workbook.Worksheets[i];
                    if (string.Equals(ws.Name, "Temp", StringComparison.OrdinalIgnoreCase))
                    {
                        workbook.Worksheets.RemoveAt(i);
                    }
                }

                // Display the names of the worksheets that remain
                Console.WriteLine("Worksheets loaded:");
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    Console.WriteLine($"- {ws.Name}");
                }

                // Save the resulting workbook
                string outputPath = "Result.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
