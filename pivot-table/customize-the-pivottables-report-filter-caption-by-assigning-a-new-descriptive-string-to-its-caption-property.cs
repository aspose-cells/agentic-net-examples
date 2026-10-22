// Title: Set a custom caption for a PivotTable report filter (page field) using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing .xlsx file with Aspose.Cells, locates the first PivotTable, changes the Name of its first PageField to a specified caption, and saves the workbook. | Demonstrate how to programmatically rename a PivotTable report filter label in a workbook using Aspose.Cells for .NET, handling missing files and absent page fields.
// Common Searches: Aspose.Cells C# change pivot table page field label | how to rename report filter caption in Excel using Aspose.Cells | set custom name for pivot table filter field programmatically .NET | update pivot table filter caption with Aspose.Cells library | C# example to modify pivot table report filter name in existing workbook
// Tags: Aspose.Cells set pivot table page field caption | C# rename pivot table report filter label | modify pivot table filter name programmatically | update pivot table filter caption .NET | custom pivot filter caption Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an existing Excel workbook, checks for a PivotTable, accesses its first report filter (page field), assigns a new string to the field's Name property to serve as a custom caption, and saves the modified workbook, with error handling for missing files and absent filters.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Access the first PivotTable
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Check for report filter (page) fields
            if (pivotTable.PageFields.Count > 0)
            {
                // Retrieve the first report filter field
                PivotField reportFilter = pivotTable.PageFields[0];

                try
                {
                    // Attempt to set a custom caption if the API supports it
                    // The CustomName property may not be available in older versions,
                    // so we fall back to setting the Name property as a best effort.
                    // reportFilter.CustomName = "Custom Filter Caption"; // Uncomment if supported
                    reportFilter.Name = "Custom Filter Caption";
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unable to set custom caption: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("PivotTable has no report filter (page) fields.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
