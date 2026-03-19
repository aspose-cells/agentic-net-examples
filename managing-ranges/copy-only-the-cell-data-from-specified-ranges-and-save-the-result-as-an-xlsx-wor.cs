using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsCopyDataExample
{
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with your actual file path)
            Workbook sourceWorkbook = new Workbook("source.xlsx");

            // Create a new empty workbook for the result
            Workbook destinationWorkbook = new Workbook();

            // Get the first worksheet from each workbook (adjust index if needed)
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

            // Define the ranges you want to copy (you can add more as needed)
            string[] sourceRangeAddresses = { "A1:B5", "D10:E12" };
            string[] destinationRangeAddresses = { "A1:B5", "D10:E12" };

            // Loop through each pair of source/destination ranges and copy only the cell data
            for (int i = 0; i < sourceRangeAddresses.Length; i++)
            {
                // Create source and destination Range objects
                AsposeRange srcRange = sourceSheet.Cells.CreateRange(sourceRangeAddresses[i]);
                AsposeRange destRange = destinationSheet.Cells.CreateRange(destinationRangeAddresses[i]);

                // Copy only the cell data (values and formulas) from source to destination
                destRange.CopyData(srcRange);
            }

            // Save the resulting workbook as XLSX
            destinationWorkbook.Save("result.xlsx");
        }
    }
}