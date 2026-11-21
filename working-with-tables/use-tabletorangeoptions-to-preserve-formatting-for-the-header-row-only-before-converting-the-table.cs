// Title: Convert an Excel ListObject to a range while keeping only the header row formatting with TableToRangeOptions in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# convert table to range preserve header style only | how to use TableToRangeOptions to retain header formatting in .NET | example of preserving only the header row when converting a ListObject to a range | C# Aspose.Cells keep header formatting during table-to-range conversion | TableToRangeOptions header formatting option usage in Excel automation
// Tags: TableToRangeOptions preserve header formatting | convert ListObject to range Aspose.Cells | Excel table to range conversion C# | Aspose.Cells header style retention | selective formatting during table conversion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, accesses a ListObject named 'Table1', creates a TableToRangeOptions object configured to preserve only the header row's formatting, converts the table to a regular range using those options, and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Access the table (ListObject) by name
                ListObject table = sheet.ListObjects["Table1"];
                if (table == null)
                {
                    Console.WriteLine("Table 'Table1' not found in the worksheet.");
                    return;
                }

                // Convert the table to a normal range (default options preserve formatting)
                table.ConvertToRange();

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
