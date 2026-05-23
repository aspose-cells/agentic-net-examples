using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeToCsv
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate some data, including commas inside the text
                cells["A1"].PutValue("Name");
                cells["B1"].PutValue("Description");
                cells["A2"].PutValue("Item1");
                cells["B2"].PutValue("Contains, comma");
                cells["A3"].PutValue("Item2");
                cells["B3"].PutValue("No comma");
                cells["A4"].PutValue("Item3");
                cells["B4"].PutValue("Another, example, with commas");

                // Define the range that we want to export (A1:B4)
                AsposeRange exportRange = cells.CreateRange("A1:B4");

                // Configure text save options for CSV export
                TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
                {
                    // Use comma as the separator
                    Separator = ',',
                    // Quote values only when needed (values containing commas will be quoted)
                    QuoteType = TxtValueQuoteType.Minimum,
                    // Specify the exact area to export
                    ExportArea = new CellArea
                    {
                        StartRow = exportRange.FirstRow,
                        StartColumn = exportRange.FirstColumn,
                        EndRow = exportRange.FirstRow + exportRange.RowCount - 1,
                        EndColumn = exportRange.FirstColumn + exportRange.ColumnCount - 1
                    }
                };

                // Ensure the output directory exists
                string outputPath = "ExportedRange.csv";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the selected range to a CSV file; commas inside text are properly escaped
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"CSV file saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}