// Title: Convert an Excel ListObject to a normal range while preserving only the header row style using TableToRangeOptions in Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a workbook, finds the first ListObject, and converts it to a regular range using TableToRangeOptions configured to keep only the header row formatting. | Show how to apply TableToRangeOptions in Aspose.Cells so that the header row retains its style after a table‑to‑range conversion.
// Common Searches: Aspose.Cells C# preserve header formatting when converting table to range | TableToRangeOptions keep only header style Excel ListObject conversion | how to retain Excel table header style after converting to range with Aspose.Cells | convert ListObject to range without losing header formatting using Aspose.Cells | C# Aspose.Cells TableToRangeOptions example for header row only
// Tags: TableToRangeOptions preserve header formatting | convert ListObject to range Aspose.Cells | Aspose.Cells header style retention | Excel table to range conversion C# | Aspose.Cells TableToRangeOptions example

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The sample loads an existing workbook, checks that the first worksheet contains a ListObject, and then converts that table to a normal range. By configuring TableToRangeOptions to retain only the header row's formatting, the header style is preserved while the rest of the table loses its table-specific formatting. The modified workbook is saved to the specified output file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (worksheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables (ListObjects) found in the worksheet.");
                return;
            }

            // Retrieve the first table
            ListObject table = worksheet.ListObjects[0];

            // Convert the table to a normal range (default conversion options)
            table.ConvertToRange();

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
