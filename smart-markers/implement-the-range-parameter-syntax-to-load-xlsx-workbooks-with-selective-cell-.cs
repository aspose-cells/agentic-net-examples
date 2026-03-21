using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeExtractionDemo
{
    class Program
    {
        static void Main()
        {
            // Load the source XLSX workbook
            Workbook sourceWorkbook = new Workbook("source.xlsx");

            // Access the first worksheet
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Define the cell range you want to extract (e.g., B2:D5)
            string rangeAddress = "B2:D5";

            // Create a Range object representing the specified area
            AsposeRange extractedRange = sourceSheet.Cells.CreateRange(rangeAddress);

            // Create a new workbook to hold the extracted range
            Workbook targetWorkbook = new Workbook();

            // Access the first worksheet in the target workbook and rename it
            Worksheet targetSheet = targetWorkbook.Worksheets[0];
            targetSheet.Name = "ExtractedRange";

            // Copy the extracted range into the target worksheet starting at cell A1
            AsposeRange destinationRange = targetSheet.Cells.CreateRange(0, 0, extractedRange.RowCount, extractedRange.ColumnCount);
            destinationRange.Copy(extractedRange);

            // Save the target workbook containing only the selected range
            targetWorkbook.Save("extracted_range.xlsx");

            Console.WriteLine("Range extraction completed successfully.");
        }
    }
}