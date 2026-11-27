// Title: Remove all shapes from each worksheet and save the workbook as CSV using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# delete all worksheet shapes before exporting to CSV | How to clear the Shapes collection in an Excel workbook with Aspose.Cells and save as CSV | Remove drawings from Excel sheets using Aspose.Cells and generate CSV output | C# code to strip images and charts from worksheets then convert workbook to CSV with Aspose.Cells
// Tags: clear worksheet shapes Aspose.Cells | save workbook as CSV Aspose.Cells | remove drawings prior to CSV export | Shapes.Clear API Aspose.Cells | CSV conversion without images Aspose.Cells

using System;
using Aspose.Cells;

namespace RemoveShapesAndExportCsv
{
    // The example loads an Excel workbook, iterates through all worksheets to clear their Shapes collections using the Shapes.Clear method, and then saves the cleaned workbook directly to a CSV file.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate arguments
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: RemoveShapesAndExportCsv <inputWorkbookPath> <outputCsvPath>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            try
            {
                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets and clear their shapes
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // The Shapes collection provides a Clear method that removes all shapes
                    sheet.Shapes.Clear();
                }

                // Save the cleaned workbook as CSV
                // SaveFormat.Csv ensures the file is written in CSV format
                workbook.Save(outputPath, SaveFormat.Csv);

                Console.WriteLine($"Successfully removed all shapes and saved CSV to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
