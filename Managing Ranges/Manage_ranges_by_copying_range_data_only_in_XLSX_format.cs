using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopyDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (must be XLSX)
            Workbook workbook = new Workbook("source.xlsx");

            // Get the first worksheet as the source sheet
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Add a new worksheet to serve as the destination
            int destIndex = workbook.Worksheets.Add();
            Worksheet destSheet = workbook.Worksheets[destIndex];

            // Define the source range (e.g., A1:C5) on the source sheet
            AsposeRange sourceRange = sourceSheet.Cells.CreateRange("A1:C5");

            // Define the destination range (e.g., E1:G5) on the destination sheet
            AsposeRange destRange = destSheet.Cells.CreateRange("E1:G5");

            // Copy only the cell data (values and formulas) from source to destination
            destRange.CopyData(sourceRange);

            // Save the workbook in XLSX format
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}