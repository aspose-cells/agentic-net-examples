using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class HideColumnUsingEntireColumn
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Create a range that starts at cell B1 (row 0, column 1)
                AsposeRange range = cells.CreateRange(0, 1, 1, 1);

                // Obtain the Range object that represents the entire column B
                AsposeRange entireColumn = range.EntireColumn;

                // Hide column B by setting its IsHidden property.
                // Column indexes are zero‑based, so column B is index 1.
                cells.Columns[1].IsHidden = true;

                // Define output file path
                string outputPath = "HideColumnB.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}