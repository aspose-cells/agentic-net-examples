using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "Report.xlsx";
                const string outputPath = "ReportUpdated.xlsx";

                // Ensure the source workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the named range "MyRange"
                Name namedRange = workbook.Worksheets.Names["MyRange"];
                if (namedRange == null)
                {
                    Console.WriteLine("Named range 'MyRange' not found.");
                    return;
                }

                // Get the actual range object (fully qualified to avoid ambiguity)
                Aspose.Cells.Range range = namedRange.GetRange();

                // Update the first cell of the range
                range.Worksheet.Cells[range.FirstRow, range.FirstColumn].PutValue("UpdatedValue");

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
}