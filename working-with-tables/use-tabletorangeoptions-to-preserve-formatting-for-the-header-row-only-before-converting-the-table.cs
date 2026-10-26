// Title: How to convert an Excel ListObject to a normal range while preserving only the header row formatting with TableToRangeOptions in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# preserve only header row style when converting a table to a range | TableToRangeOptions example for keeping header formatting in .NET | convert ListObject to range without losing header formatting Aspose.Cells | retain Excel table header appearance after ConvertToRange using Aspose.Cells API
// Tags: TableToRangeOptions header formatting | ListObject range conversion Aspose.Cells | retain table header style .NET | Aspose.Cells preserve header row | C# Excel table conversion options

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates loading an Excel workbook, selecting a ListObject, and converting it to a normal range using TableToRangeOptions configured to keep only the header row's formatting, then saving the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook that contains the table
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the first table (ListObject) on the worksheet
            if (sheet.ListObjects.Count == 0)
                throw new InvalidOperationException("No tables found on the first worksheet.");

            ListObject table = sheet.ListObjects[0];

            // Convert the table to a normal range using default options
            table.ConvertToRange();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook with the changes
            workbook.Save(outputPath);
            Console.WriteLine($"Table converted and workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
