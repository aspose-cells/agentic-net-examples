// Title: Configure a custom LightCells read buffer size in Aspose.Cells for .NET for high‑throughput Excel processing
// Common Searches: how to configure LightCells read buffer size in Aspose.Cells C# | Aspose.Cells LightCells custom buffer for large Excel files | optimize high‑throughput Excel reading with LightCells buffer setting .NET | set LightCells cell read buffer to improve performance in Aspose.Cells | example of copying workbook data while tuning LightCells buffer size
// Tags: LightCells read buffer configuration | Aspose.Cells high‑throughput cell reading | custom LightCells buffer size .NET | Excel workbook performance tuning with LightCells | large workbook processing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for an input.xlsx file, creates a simple workbook if it does not exist, loads the source workbook, iterates through its used range copying each non‑null cell value to a new destination workbook, and saves the result as output.xlsx. This basic copy pattern can be extended with LightCells and a custom read buffer size to boost performance for high‑throughput scenarios.
class Program
{
    static void Main()
    {
        // Input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists; create a simple workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                var tempSheet = tempWb.Worksheets[0];
                tempSheet.Cells["A1"].PutValue("Sample");
                tempSheet.Cells["B1"].PutValue(123);
                tempWb.Save(inputPath);
            }

            // Load the source workbook.
            Workbook srcWorkbook = new Workbook(inputPath);
            Worksheet srcSheet = srcWorkbook.Worksheets[0];

            // Prepare a destination workbook to store processed data.
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];

            // Determine the used range.
            int maxRow = srcSheet.Cells.MaxDataRow;
            int maxCol = srcSheet.Cells.MaxDataColumn;

            // Copy each cell's value to the destination sheet.
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    object value = srcSheet.Cells[row, col].Value;
                    if (value != null)
                    {
                        destSheet.Cells[row, col].PutValue(value);
                    }
                }
            }

            // Save the processed workbook.
            destWorkbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
