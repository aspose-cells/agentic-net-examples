// Title: Programmatically set an Excel table's style to TableStyleLight10 using Aspose.Cells for .NET
// AI Prompts: Set the TableStyleName of the first ListObject in a worksheet to "TableStyleLight10" and save the workbook. | Loop through all tables (ListObjects) in a workbook and apply the built‑in "TableStyleLight10" style with Aspose.Cells for .NET. | Load an existing Excel file, verify that it contains tables, change their style to TableStyleLight10, and handle missing files or style errors gracefully.
// Common Searches: Aspose.Cells C# set Excel table style to TableStyleLight10 | How to apply a built‑in table style to a ListObject using Aspose.Cells .NET | Change Excel table formatting programmatically with Aspose.Cells for .NET
// Tags: Aspose.Cells apply TableStyleLight10 to ListObject | C# set Excel table style programmatically | Aspose.Cells built‑in table formatting | Excel workbook table style matching color palette .NET | ListObject TableStyleName property usage

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject

// The example loads a workbook, checks for tables on the first worksheet, sets the first table's TableStyleName to "TableStyleLight10" (a built‑in style that aligns with the workbook's palette), and saves the result, with error handling for missing files and style‑application failures.
class Program
{
    static void Main()
    {
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

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or change the index as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Check if the worksheet contains at least one table (ListObject)
            if (worksheet.ListObjects.Count > 0)
            {
                // Retrieve the first table in the worksheet
                ListObject table = worksheet.ListObjects[0];

                // Apply a built‑in table style using the TableStyleName property
                try
                {
                    // Aspose.Cells allows setting the style by name (e.g., "TableStyleLight10")
                    table.TableStyleName = "TableStyleLight10";
                }
                catch (Exception ex)
                {
                    // Log but continue if style application fails
                    Console.WriteLine($"Table style could not be applied: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("No tables found in the worksheet.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook to the desired output path
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
