using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        try
        {
            // Create the source workbook and populate it with sample data
            Workbook srcWorkbook = new Workbook();
            Worksheet srcSheet = srcWorkbook.Worksheets[0];

            for (int row = 0; row < 5; row++)
            {
                srcSheet.Cells[row, 0].PutValue($"Item {row + 1}");
                srcSheet.Cells[row, 1].PutValue(row * 10);
                srcSheet.Cells[row, 2].PutValue(DateTime.Today.AddDays(row));
            }

            // Define the source range (A1:C5)
            AsposeRange sourceRange = srcSheet.Cells.CreateRange(0, 0, 5, 3);

            // Create the destination workbook where the range will be copied
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];
            AsposeRange destRange = destSheet.Cells.CreateRange(0, 0, 5, 3);

            // Copy the source range to the destination range
            sourceRange.Copy(destRange);

            // Convert the copied range into a table (ListObject) to get filtering/sorting
            int tableIdx = destSheet.ListObjects.Add(0, 0, 4, 2, true);
            ListObject table = destSheet.ListObjects[tableIdx];
            table.ApplyStyleToRange(); // applies default style and adds auto‑filter

            // Save the resulting workbook
            string outputPath = "CopiedRangeWithTable.xlsx";
            destWorkbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}