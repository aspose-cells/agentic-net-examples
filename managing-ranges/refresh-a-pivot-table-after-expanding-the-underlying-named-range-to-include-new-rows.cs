using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class RefreshPivotAfterRangeExpansion
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing the named range and pivot table
            Workbook workbook = new Workbook(inputPath);

            // Assume the data source is on the first worksheet
            Worksheet dataSheet = workbook.Worksheets[0];

            // Append new rows of data below the existing data
            int startRow = dataSheet.Cells.MaxDataRow + 1; // first empty row after existing data
            dataSheet.Cells[startRow, 0].PutValue("NewItem1");   // Column A
            dataSheet.Cells[startRow, 1].PutValue(123);         // Column B
            dataSheet.Cells[startRow + 1, 0].PutValue("NewItem2");
            dataSheet.Cells[startRow + 1, 1].PutValue(456);

            // Expand the named range "DataRange" to include the newly added rows
            // Retrieve the existing named range from the workbook's Names collection
            Name dataRange = workbook.Worksheets.Names["DataRange"];
            if (dataRange != null)
            {
                // Build a new address that covers the original range plus the new rows
                string sheetName = dataSheet.Name;
                string newAddress = $"={sheetName}!$A$1:$B${startRow + 1}";
                dataRange.RefersTo = newAddress;
            }
            else
            {
                Console.WriteLine("Named range 'DataRange' not found.");
            }

            // Refresh pivot tables so they pick up the expanded range
            // Assuming the pivot table is on the same sheet; adjust index if needed
            Worksheet pivotSheet = workbook.Worksheets[0];
            pivotSheet.RefreshPivotTables();

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}