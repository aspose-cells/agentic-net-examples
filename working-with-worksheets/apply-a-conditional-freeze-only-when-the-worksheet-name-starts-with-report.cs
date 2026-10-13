// Title: Freeze the first row on Excel worksheets whose names start with “Report” using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, loops through all worksheets, and calls Worksheet.FreezePanes to freeze the top row only when the sheet name begins with "Report". | Show how to perform a case‑insensitive check on a worksheet's Name property and apply FreezePanes conditionally in Aspose.Cells for .NET. | Provide a complete example that verifies the input file, applies selective row freezing, and saves the updated workbook.
// Common Searches: asp.net freeze first row only on sheets starting with report | aspocells conditional FreezePanes based on worksheet name | c# iterate worksheets and apply freeze panes if name starts with Report | how to freeze panes selectively in an Excel workbook using Aspose.Cells | apply FreezePanes to specific worksheets in .NET
// Tags: conditional Worksheet.FreezePanes Aspose.Cells | freeze top row on sheets with Report prefix | C# iterate workbook worksheets Aspose.Cells | selective freeze panes Excel .NET | Aspose.Cells load and save workbook example

using System;
using System.IO;
using Aspose.Cells;

// The sample loads input.xlsx, checks each worksheet, and uses Worksheet.FreezePanes(1,0,1,0) to freeze the first row only on sheets whose names start with "Report", then saves the modified workbook to output.xlsx.
class ConditionalFreezeExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Apply freeze only when the worksheet name starts with "Report"
                if (sheet.Name.StartsWith("Report", StringComparison.OrdinalIgnoreCase))
                {
                    // Freeze the first row (row index is zero‑based)
                    // Parameters: top‑left cell of scrollable area (row, column) and number of rows/columns to freeze
                    sheet.FreezePanes(1, 0, 1, 0);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
